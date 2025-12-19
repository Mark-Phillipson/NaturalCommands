# Show Letters Feature - Testing Guide

This document describes how to test the new "show letters" feature for voice-based UI element navigation.

## Feature Overview

The "show letters" feature overlays two-letter labels (a, b, c, ..., aa, ab, ...) on all clickable UI elements in the active window, allowing users to quickly navigate and click elements using voice commands via Talon.

## Prerequisites

- Windows 10 or later
- .NET 10.0 runtime
- Built NaturalCommands application

## Manual Testing Steps

### Test 1: Basic Functionality with Notepad

1. Open Notepad
2. Run: `NaturalCommands.exe natural "show letters"`
3. **Expected Result**: 
   - A transparent overlay appears
   - Letter labels (a, b, c, etc.) are shown on buttons like "File", "Edit", "Format", etc.
   - Instructions are displayed at the bottom: "Type letters to click elements. Press ESC to cancel."

4. Type a letter (e.g., "a") corresponding to a button
5. **Expected Result**: 
   - The corresponding button/menu item is clicked
   - The overlay closes automatically

6. Run the command again
7. Press ESC
8. **Expected Result**: The overlay closes without clicking anything

### Test 2: Web Browser (Microsoft Edge)

1. Open Microsoft Edge with any webpage
2. Run: `NaturalCommands.exe natural "show letters"`
3. **Expected Result**: 
   - Labels appear on all clickable elements (links, buttons, form controls)
   - Labels are positioned at the top-left of each element

4. Type letters to click a link
5. **Expected Result**: The link is clicked and the page navigates

### Test 3: Multi-letter Labels

1. Open an application with many UI elements (e.g., Visual Studio, Microsoft Office)
2. Run: `NaturalCommands.exe natural "show letters"`
3. **Expected Result**: 
   - Single letters are used first (a-z)
   - Two-letter combinations are used when needed (aa, ab, ac, ...)
   - All clickable elements have unique labels

4. Type a two-letter combination (e.g., "ab")
5. **Expected Result**: The corresponding element is clicked

### Test 4: Backspace Handling

1. Run the show letters command
2. Type a letter (e.g., "a")
3. Press Backspace
4. Type a different letter (e.g., "b")
5. **Expected Result**: The second element is clicked (backspace cleared the first letter)

### Test 5: Invalid Label

1. Run the show letters command
2. Type two letters that don't match any label (e.g., "zz" if no such label exists)
3. **Expected Result**: Nothing happens, the typed text is cleared, and the overlay remains open

### Test 6: Different Control Types

Test with an application that has various control types:

1. **Buttons**: Regular push buttons
2. **Links**: Hyperlinks in browsers
3. **Checkboxes**: Toggle controls
4. **Radio buttons**: Selection controls
5. **Menu items**: Dropdown menu entries
6. **Tab items**: Tab controls
7. **Combo boxes**: Dropdown lists

**Expected Result**: All these control types should be labeled and clickable

### Test 7: Multiple Monitors

1. If you have multiple monitors, test with applications on different screens
2. Run: `NaturalCommands.exe natural "show letters"`
3. **Expected Result**: 
   - The overlay covers all screens
   - Labels are shown on the active window elements regardless of which monitor it's on

### Test 8: Voice Integration (with Talon)

If you have Talon Voice configured:

1. Say: "natural show letters"
2. **Expected Result**: The overlay appears

3. Say: "air bat" (or whatever letters correspond to an element)
4. **Expected Result**: The element is clicked

## Known Limitations

1. **UI Automation Support**: The feature only works with applications that support Windows UI Automation API. Some legacy applications may not be fully supported.

2. **Browser Limitations**: In browsers, the feature can detect buttons and top-level links, but may not detect all clickable elements within complex web applications.

3. **Performance**: With applications that have hundreds of UI elements, the initial enumeration may take 1-2 seconds.

4. **Maximum Elements**: The feature is limited to 500 elements to prevent performance issues.

## Troubleshooting

### No elements found

- **Cause**: The active window doesn't support UI Automation or has no clickable elements
- **Solution**: Try with a different application or window

### Labels overlap

- **Cause**: Elements are very close together
- **Solution**: This is expected behavior. The labels are positioned at the top-left of each element

### Click doesn't work

- **Cause**: The element doesn't support the Invoke or Toggle pattern
- **Solution**: This is a limitation of the UI Automation API. Some custom controls may not be clickable via automation

### Overlay doesn't close after clicking

- **Cause**: The click action failed
- **Solution**: Press ESC to close manually and try a different element

## Reporting Issues

If you encounter issues:

1. Check the log file at `bin/bin/app.log` for error messages
2. Note which application you were using
3. Describe the expected vs actual behavior
4. Include any error messages from the log

## Future Enhancements

Potential improvements for future versions:

1. Support for clicking elements by name/text search
2. Filter by element type (e.g., show only buttons or only links)
3. Visual feedback when typing letters (highlight matching elements)
4. Support for scrolling to off-screen elements
5. Configurable label colors and fonts
6. Persistent labels that don't close after clicking (for multiple selections)
