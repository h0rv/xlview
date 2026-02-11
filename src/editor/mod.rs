//! Cell editing wrapper around `XlView`.
//!
//! `XlEdit` wraps the read-only viewer, adding:
//! - Cell editing via an HTML input overlay
//! - Dirty tracking for modified cells
//! - XLSX save (roundtrip via the export pipeline)

#[cfg(target_arch = "wasm32")]
mod input;
pub(crate) mod mutation;

use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use wasm_bindgen::prelude::*;

use crate::viewer::XlView;

#[cfg(target_arch = "wasm32")]
use input::InputOverlay;

#[cfg(target_arch = "wasm32")]
use web_sys::HtmlCanvasElement;

/// A cell edit record.
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub(crate) struct CellEdit {
    /// The new value as entered by the user.
    pub value: String,
}

/// Editor state (separate from viewer's SharedState).
pub(crate) struct EditorState {
    /// Currently editing cell `(row, col)`, or `None`.
    pub editing_cell: Option<(u32, u32)>,
    /// Edits keyed by `(sheet_idx, row, col)`.
    pub dirty_cells: HashMap<(usize, u32, u32), CellEdit>,
    /// Set of sheet indices that have been modified.
    pub dirty_sheets: HashSet<usize>,
}

/// The main editor struct exported to JavaScript.
///
/// Wraps `XlView` (read-only viewer) and adds editing + save capabilities.
#[wasm_bindgen]
pub struct XlEdit {
    #[cfg(target_arch = "wasm32")]
    viewer: XlView,
    #[cfg(target_arch = "wasm32")]
    original_bytes: Option<Vec<u8>>,
    #[cfg(target_arch = "wasm32")]
    editor_state: Rc<RefCell<EditorState>>,
    #[cfg(target_arch = "wasm32")]
    input_overlay: InputOverlay,

    // Non-wasm32 fields (for tests/CLI)
    #[cfg(not(target_arch = "wasm32"))]
    viewer: XlView,
    #[cfg(not(target_arch = "wasm32"))]
    original_bytes: Option<Vec<u8>>,
    #[cfg(not(target_arch = "wasm32"))]
    editor_state: Rc<RefCell<EditorState>>,
}

// ============================================================================
// WASM32 Implementation
// ============================================================================

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen]
impl XlEdit {
    /// Create a new editor wrapping an XlView.
    #[wasm_bindgen(constructor)]
    pub fn new(
        base_canvas: HtmlCanvasElement,
        overlay_canvas: HtmlCanvasElement,
        dpr: f32,
    ) -> Result<XlEdit, JsValue> {
        let viewer = XlView::new_with_overlay(base_canvas, overlay_canvas, dpr)?;
        let editor_state = Rc::new(RefCell::new(EditorState {
            editing_cell: None,
            dirty_cells: HashMap::new(),
            dirty_sheets: HashSet::new(),
        }));
        let input_overlay = InputOverlay::new();

        Ok(XlEdit {
            viewer,
            original_bytes: None,
            editor_state,
            input_overlay,
        })
    }

    /// Load an XLSX file from bytes.
    ///
    /// Stores the original bytes for later save (ZIP roundtrip).
    #[wasm_bindgen]
    pub fn load(&mut self, data: &[u8]) -> Result<(), JsValue> {
        self.original_bytes = Some(data.to_vec());
        // Clear edit state
        {
            let mut es = self.editor_state.borrow_mut();
            es.editing_cell = None;
            es.dirty_cells.clear();
            es.dirty_sheets.clear();
        }
        self.input_overlay.hide();
        self.viewer.load(data)
    }

    /// Begin editing a cell at `(row, col)`.
    ///
    /// Shows an input overlay positioned over the cell.
    #[wasm_bindgen]
    pub fn begin_edit(&mut self, row: u32, col: u32) {
        // Get current cell value
        let current_value = self.viewer.cell_value(row, col).unwrap_or_default();

        // Get cell rect for positioning
        let Some(rect) = self.viewer.cell_rect(row, col) else {
            return;
        };

        {
            let mut es = self.editor_state.borrow_mut();
            es.editing_cell = Some((row, col));
        }

        // Get the scroll container to position input relative to it
        let container = self.viewer.scroll_container_element();

        self.input_overlay
            .show(&rect, &current_value, container.as_ref());
    }

    /// Commit the current edit with the given value.
    #[wasm_bindgen]
    pub fn commit_edit(&mut self, value: &str) -> Result<(), JsValue> {
        let (row, col, sheet_idx) = {
            let es = self.editor_state.borrow();
            let Some((row, col)) = es.editing_cell else {
                return Ok(());
            };
            let state = self.viewer.shared_state();
            let s = state.borrow();
            (row, col, s.active_sheet)
        };

        // Apply edit to the workbook
        {
            let state = self.viewer.shared_state();
            let mut s = state.borrow_mut();
            if let Some(ref mut workbook) = s.workbook {
                mutation::apply_cell_edit(workbook, sheet_idx, row, col, value)
                    .map_err(|e| JsValue::from_str(&e.to_string()))?;
            }
        }

        // Track the edit
        {
            let mut es = self.editor_state.borrow_mut();
            es.editing_cell = None;
            es.dirty_cells.insert(
                (sheet_idx, row, col),
                CellEdit {
                    value: value.to_string(),
                },
            );
            es.dirty_sheets.insert(sheet_idx);
        }

        // Hide input and trigger re-render (clear tile cache so edits are visible)
        self.input_overlay.hide();
        self.viewer.invalidate_data();

        Ok(())
    }

    /// Cancel the current edit.
    #[wasm_bindgen]
    pub fn cancel_edit(&mut self) {
        {
            let mut es = self.editor_state.borrow_mut();
            es.editing_cell = None;
        }
        self.input_overlay.hide();
    }

    /// Get the current value from the input overlay.
    #[wasm_bindgen]
    pub fn input_value(&self) -> Option<String> {
        self.input_overlay.value()
    }

    /// Save the workbook to XLSX bytes.
    ///
    /// Returns the modified XLSX as a `Vec<u8>`. If nothing was edited,
    /// returns the original bytes.
    #[wasm_bindgen]
    pub fn save(&self) -> Result<Vec<u8>, JsValue> {
        let Some(ref original) = self.original_bytes else {
            return Err(JsValue::from_str("no file loaded"));
        };

        let es = self.editor_state.borrow();

        if es.dirty_sheets.is_empty() {
            return Ok(original.clone());
        }

        let state = self.viewer.shared_state();
        let s = state.borrow();
        let Some(ref workbook) = s.workbook else {
            return Err(JsValue::from_str("no workbook loaded"));
        };

        crate::export::save_xlsx(original, workbook, &es.dirty_sheets)
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Check if any edits have been made.
    #[wasm_bindgen]
    pub fn is_dirty(&self) -> bool {
        let es = self.editor_state.borrow();
        !es.dirty_sheets.is_empty()
    }

    /// Check if currently editing a cell.
    #[wasm_bindgen]
    pub fn is_editing(&self) -> bool {
        let es = self.editor_state.borrow();
        es.editing_cell.is_some()
    }

    // ---- Delegate methods to the inner viewer ----

    /// Render the current state.
    #[wasm_bindgen]
    pub fn render(&mut self) -> Result<(), JsValue> {
        self.viewer.render()
    }

    /// Resize the viewport.
    #[wasm_bindgen]
    pub fn resize(&mut self, physical_width: u32, physical_height: u32, dpr: f32) {
        self.viewer.resize(physical_width, physical_height, dpr);
    }

    /// Switch to a different sheet.
    #[wasm_bindgen]
    pub fn set_active_sheet(&mut self, index: usize) {
        self.cancel_edit();
        self.viewer.set_active_sheet(index);
    }

    /// Force a re-render.
    #[wasm_bindgen]
    pub fn invalidate(&mut self) {
        self.viewer.invalidate();
    }

    /// Get the current selection.
    #[wasm_bindgen]
    pub fn get_selection(&self) -> Option<Vec<u32>> {
        self.viewer.get_selection()
    }

    /// Register a render callback.
    #[wasm_bindgen]
    pub fn set_render_callback(&mut self, callback: Option<js_sys::Function>) {
        self.viewer.set_render_callback(callback);
    }

    /// Get sheet names.
    #[wasm_bindgen]
    pub fn sheet_names(&self) -> Vec<String> {
        self.viewer.sheet_names()
    }

    /// Get active sheet index.
    #[wasm_bindgen]
    pub fn active_sheet(&self) -> usize {
        self.viewer.active_sheet()
    }

    /// Hit-test: which cell is at the given viewport point?
    #[wasm_bindgen]
    pub fn cell_at_point(&self, x: f32, y: f32) -> Option<Vec<u32>> {
        self.viewer.cell_at_point(x, y)
    }
}

// ============================================================================
// Non-WASM32 Implementation (for tests)
// ============================================================================

#[cfg(not(target_arch = "wasm32"))]
impl XlEdit {
    /// Create a new editor (non-WASM, for testing/CLI).
    #[must_use]
    pub fn new_test() -> Self {
        let viewer = XlView::new_test(800, 600, 1.0);
        let editor_state = Rc::new(RefCell::new(EditorState {
            editing_cell: None,
            dirty_cells: HashMap::new(),
            dirty_sheets: HashSet::new(),
        }));
        XlEdit {
            viewer,
            original_bytes: None,
            editor_state,
        }
    }

    /// Load an XLSX file from bytes.
    pub fn load(&mut self, data: &[u8]) -> crate::error::Result<()> {
        self.original_bytes = Some(data.to_vec());
        {
            let mut es = self.editor_state.borrow_mut();
            es.editing_cell = None;
            es.dirty_cells.clear();
            es.dirty_sheets.clear();
        }
        self.viewer.load(data)
    }

    /// Commit an edit to a cell.
    pub fn commit_edit(
        &mut self,
        sheet_idx: usize,
        row: u32,
        col: u32,
        value: &str,
    ) -> crate::error::Result<()> {
        // Apply edit directly to the viewer's workbook
        if let Some(workbook) = self.viewer.workbook_mut() {
            mutation::apply_cell_edit(workbook, sheet_idx, row, col, value)?;
        }

        let mut es = self.editor_state.borrow_mut();
        es.dirty_cells.insert(
            (sheet_idx, row, col),
            CellEdit {
                value: value.to_string(),
            },
        );
        es.dirty_sheets.insert(sheet_idx);
        Ok(())
    }

    /// Save the workbook to XLSX bytes.
    pub fn save(&self) -> crate::error::Result<Vec<u8>> {
        let original = self
            .original_bytes
            .as_ref()
            .ok_or_else(|| crate::error::XlviewError::Other("no file loaded".into()))?;

        let es = self.editor_state.borrow();
        if es.dirty_sheets.is_empty() {
            return Ok(original.clone());
        }

        let workbook = self
            .viewer
            .workbook_ref()
            .ok_or_else(|| crate::error::XlviewError::Other("no workbook loaded".into()))?;

        crate::export::save_xlsx(original, workbook, &es.dirty_sheets)
    }

    /// Check if any edits have been made.
    pub fn is_dirty(&self) -> bool {
        let es = self.editor_state.borrow();
        !es.dirty_sheets.is_empty()
    }
}
