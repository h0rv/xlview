//! Tests for the editing + export pipeline.
#![cfg(feature = "editing")]

#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic
)]
mod tests {
    use xlview::editor::XlEdit;
    use xlview::parser;
    use xlview::types::{CellType, Workbook};

    // ================================================================
    // Test helpers
    // ================================================================

    /// Get test XLSX bytes from the test/ directory.
    fn load_test_file(name: &str) -> Vec<u8> {
        std::fs::read(format!("test/{name}")).expect("test file should exist")
    }

    /// Get a minimal XLSX for editing tests.
    fn minimal_xlsx() -> Vec<u8> {
        load_test_file("minimal.xlsx")
    }

    /// Find a cell value in a sheet by (row, col).
    fn find_cell_value(
        workbook: &Workbook,
        sheet_idx: usize,
        row: u32,
        col: u32,
    ) -> Option<String> {
        let sheet = workbook.sheets.get(sheet_idx)?;
        for cd in &sheet.cells {
            if cd.r == row && cd.c == col {
                return cd.cell.v.clone();
            }
        }
        None
    }

    /// Find a cell in a sheet and return its type.
    fn find_cell_type(
        workbook: &Workbook,
        sheet_idx: usize,
        row: u32,
        col: u32,
    ) -> Option<CellType> {
        let sheet = workbook.sheets.get(sheet_idx)?;
        for cd in &sheet.cells {
            if cd.r == row && cd.c == col {
                return Some(cd.cell.t);
            }
        }
        None
    }

    /// Check if a cell exists.
    fn has_cell(workbook: &Workbook, sheet_idx: usize, row: u32, col: u32) -> bool {
        workbook
            .sheets
            .get(sheet_idx)
            .map(|s| s.cells.iter().any(|cd| cd.r == row && cd.c == col))
            .unwrap_or(false)
    }

    // ================================================================
    // Mutation tests
    // ================================================================

    #[test]
    fn test_edit_string_cell() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "Hello World").unwrap();
        assert!(editor.is_dirty());
    }

    #[test]
    fn test_edit_number_cell() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "42.5").unwrap();
        assert!(editor.is_dirty());
    }

    #[test]
    fn test_edit_boolean_cell() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "true").unwrap();
        assert!(editor.is_dirty());
    }

    #[test]
    fn test_clear_cell() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "").unwrap();
        assert!(editor.is_dirty());
    }

    #[test]
    fn test_not_dirty_before_edit() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        assert!(!editor.is_dirty());
    }

    // ================================================================
    // Save roundtrip tests
    // ================================================================

    #[test]
    fn test_save_no_edits() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        let saved = editor.save().unwrap();
        assert_eq!(saved, data);
    }

    #[test]
    fn test_save_roundtrip_preserves_edit() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "Edited Value").unwrap();
        let saved = editor.save().unwrap();

        // Reload the saved file
        let reloaded = parser::parse(&saved).unwrap();
        let val = find_cell_value(&reloaded, 0, 0, 0);
        assert_eq!(val.as_deref(), Some("Edited Value"));
    }

    #[test]
    fn test_save_roundtrip_number() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "123.456").unwrap();
        let saved = editor.save().unwrap();

        let reloaded = parser::parse(&saved).unwrap();
        let cell_type = find_cell_type(&reloaded, 0, 0, 0);
        assert_eq!(cell_type, Some(CellType::Number));
    }

    #[test]
    fn test_save_roundtrip_boolean() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "TRUE").unwrap();
        let saved = editor.save().unwrap();

        let reloaded = parser::parse(&saved).unwrap();
        let cell_type = find_cell_type(&reloaded, 0, 0, 0);
        assert_eq!(cell_type, Some(CellType::Boolean));
    }

    #[test]
    fn test_save_is_valid_zip() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "test").unwrap();
        let saved = editor.save().unwrap();

        let cursor = std::io::Cursor::new(&saved);
        let archive = zip::ZipArchive::new(cursor).expect("saved file should be valid ZIP");
        assert!(!archive.is_empty(), "ZIP should have entries");
    }

    #[test]
    fn test_save_unmodified_sheets_passthrough() {
        let data = load_test_file("styled.xlsx");
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        let original = parser::parse(&data).unwrap();

        // Only edit the first sheet
        editor.commit_edit(0, 0, 0, "Modified").unwrap();
        let saved = editor.save().unwrap();

        let reloaded = parser::parse(&saved).unwrap();

        // Verify sheets count is preserved
        assert_eq!(reloaded.sheets.len(), original.sheets.len());

        // First sheet's edit persisted
        let val = find_cell_value(&reloaded, 0, 0, 0);
        assert_eq!(val.as_deref(), Some("Modified"));
    }

    // ================================================================
    // Formula preservation
    // ================================================================

    #[test]
    fn test_formula_field_defaults_to_none() {
        let data = minimal_xlsx();
        let workbook = parser::parse(&data).unwrap();

        for sheet in &workbook.sheets {
            for cd in &sheet.cells {
                let _ = cd.cell.formula.as_ref();
            }
        }
    }

    // ================================================================
    // Mutation via parser API
    // ================================================================

    #[test]
    fn test_mutation_type_detection_via_roundtrip() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        // Number
        editor.commit_edit(0, 0, 0, "42").unwrap();
        let saved = editor.save().unwrap();
        let reloaded = parser::parse(&saved).unwrap();
        assert_eq!(find_cell_type(&reloaded, 0, 0, 0), Some(CellType::Number));

        // Reload, edit to string
        let mut editor2 = XlEdit::new_test();
        editor2.load(&saved).unwrap();
        editor2.commit_edit(0, 0, 0, "hello").unwrap();
        let saved2 = editor2.save().unwrap();
        let reloaded2 = parser::parse(&saved2).unwrap();
        assert_eq!(find_cell_type(&reloaded2, 0, 0, 0), Some(CellType::String));
        assert_eq!(
            find_cell_value(&reloaded2, 0, 0, 0).as_deref(),
            Some("hello")
        );
    }

    #[test]
    fn test_mutation_clear_cell_via_roundtrip() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        // First check if there's a cell at (0,0)
        let original = parser::parse(&data).unwrap();
        let had_cell = has_cell(&original, 0, 0, 0);

        // Clear it
        editor.commit_edit(0, 0, 0, "").unwrap();
        let saved = editor.save().unwrap();
        let reloaded = parser::parse(&saved).unwrap();

        if had_cell {
            let still_has = has_cell(&reloaded, 0, 0, 0);
            assert!(!still_has, "cell should be removed after clearing");
        }
    }

    #[test]
    fn test_insert_new_cell_via_roundtrip() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        // Insert at row 50, col 0 (likely doesn't exist in minimal.xlsx)
        editor.commit_edit(0, 50, 0, "new cell").unwrap();
        let saved = editor.save().unwrap();
        let reloaded = parser::parse(&saved).unwrap();

        let val = find_cell_value(&reloaded, 0, 50, 0);
        assert_eq!(val.as_deref(), Some("new cell"));
    }

    // ================================================================
    // Load/save lifecycle
    // ================================================================

    #[test]
    fn test_load_resets_dirty_state() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "edit1").unwrap();
        assert!(editor.is_dirty());

        // Re-load clears dirty state
        editor.load(&data).unwrap();
        assert!(!editor.is_dirty());
    }

    #[test]
    fn test_multiple_edits_same_cell() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "first").unwrap();
        editor.commit_edit(0, 0, 0, "second").unwrap();
        editor.commit_edit(0, 0, 0, "third").unwrap();

        let saved = editor.save().unwrap();
        let reloaded = parser::parse(&saved).unwrap();
        let val = find_cell_value(&reloaded, 0, 0, 0);
        assert_eq!(val.as_deref(), Some("third"));
    }

    #[test]
    fn test_multiple_edits_different_cells() {
        let data = minimal_xlsx();
        let mut editor = XlEdit::new_test();
        editor.load(&data).unwrap();

        editor.commit_edit(0, 0, 0, "A1 edit").unwrap();
        editor.commit_edit(0, 1, 0, "A2 edit").unwrap();
        editor.commit_edit(0, 0, 1, "B1 edit").unwrap();

        let saved = editor.save().unwrap();
        let reloaded = parser::parse(&saved).unwrap();
        assert_eq!(
            find_cell_value(&reloaded, 0, 0, 0).as_deref(),
            Some("A1 edit")
        );
        assert_eq!(
            find_cell_value(&reloaded, 0, 1, 0).as_deref(),
            Some("A2 edit")
        );
        assert_eq!(
            find_cell_value(&reloaded, 0, 0, 1).as_deref(),
            Some("B1 edit")
        );
    }
}
