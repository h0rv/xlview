# BUG-009: Gradient Fills Not Parsed

**Priority**: MEDIUM
**Status**: Open
**Component**: styles.rs

## Problem

Excel supports gradient fills in cells, but the parser only handles pattern fills. Gradient fills are silently ignored, resulting in no fill being displayed.

## Current Behavior

```rust
// styles.rs - only handles patternFill
"patternFill" => {
    // ... parsing
}
// gradientFill is not handled
```

Cells with gradient fills appear as having no background.

## Expected Behavior

Parse `<gradientFill>` elements and convert to CSS gradients.

## Where Data Should Come From

```xml
<!-- xl/styles.xml -->
<fills>
  <fill>
    <gradientFill type="linear" degree="90">
      <stop position="0">
        <color rgb="FF0000FF"/>  <!-- Blue -->
      </stop>
      <stop position="1">
        <color rgb="FFFF0000"/>  <!-- Red -->
      </stop>
    </gradientFill>
  </fill>
</fills>
```

Attributes:
- `type`: "linear" or "path" (radial)
- `degree`: angle for linear gradients (0-360)
- `left`, `right`, `top`, `bottom`: for path/radial gradients
- `<stop>` elements with `position` (0-1) and color

## Impact

- Cells with gradient backgrounds appear blank
- Decorative spreadsheets lose visual appeal
- Reports with gradient headers look broken

## Proposed Fix

1. Add parsing for `<gradientFill>` element
2. Store gradient type, angle, and stops
3. Convert to CSS gradient in renderer:
   ```css
   /* Linear gradient at 90 degrees */
   background: linear-gradient(90deg, #0000FF 0%, #FF0000 100%);

   /* Radial/path gradient */
   background: radial-gradient(circle at 50% 50%, #0000FF 0%, #FF0000 100%);
   ```

## References

- ECMA-376 Part 1, Section 18.8.24 (gradientFill)
- ECMA-376 Part 1, Section 18.8.38 (stop)
