# Implementation Plan

- [x] 1. Set up history data structures




  - [x] 1.1 Add HistoryState to application state


    - Add entries array, panelVisible flag, maxEntries constant

    - _Requirements: 2.1, 3.1, 3.3_


  - [ ] 1.2 Create history.js module with core functions
    - Implement addToHistory, removeFromHistory, getHistory, clearHistory


    - _Requirements: 1.3, 2.4, 3.3_
  - [ ] 1.3 Write property test for history prepend order
    - **Property 3: New entries are prepended to history**
    - **Validates: Requirements 2.4**
  - [ ] 1.4 Write property test for max limit enforcement
    - **Property 5: History respects maximum limit**
    - **Validates: Requirements 3.3**

- [x] 2. Implement undo data capture

  - [x] 2.1 Implement captureUndoData function

    - Capture cell values and formulas before action is applied
    - _Requirements: 1.1_
  - [x] 2.2 Update handleApply to capture undo data before each action

    - Call captureUndoData before executeAction
    - _Requirements: 1.1_
  - [x] 2.3 Update handleApply to add successful actions to history

    - Create history entry with action details and undo data
    - _Requirements: 2.4_

- [ ] 3. Implement undo functionality
  - [x] 3.1 Implement restoreFromUndoData function

    - Restore cell values/formulas from captured undo data
    - _Requirements: 1.2_

  - [ ] 3.2 Implement performUndo function
    - Get most recent entry, restore data, remove from history
    - _Requirements: 1.2, 1.3_
  - [x] 3.3 Write property test for undo removes entry

    - **Property 4: Undo removes entry from history**
    - **Validates: Requirements 1.3**

- [ ] 4. Implement history panel UI
  - [x] 4.1 Add history panel HTML structure to taskpane.html

    - Add history panel container, entry list, toggle button
    - _Requirements: 2.1, 4.1_

  - [ ] 4.2 Add history panel CSS styles to taskpane.css
    - Style history entries, timestamps, toggle button
    - _Requirements: 2.2_
  - [x] 4.3 Implement formatRelativeTime function

    - Format timestamps as "X min ago", "X hours ago", etc.
    - _Requirements: 2.2_

  - [ ] 4.4 Implement renderHistoryEntry function
    - Render single entry with type icon, target, timestamp
    - _Requirements: 2.2_
  - [x] 4.5 Write property test for history entry rendering

    - **Property 2: History entry contains required fields**
    - **Validates: Requirements 2.2**

  - [ ] 4.6 Implement renderHistoryPanel function
    - Render all entries or empty state message
    - _Requirements: 2.1, 2.3_
  - [x] 4.7 Write property test for history panel rendering

    - **Property 6: History panel renders all entries**
    - **Validates: Requirements 2.1**

- [ ] 5. Integrate undo button and history toggle
  - [x] 5.1 Add Undo button to input area in taskpane.html

    - Add button next to Apply Changes
    - _Requirements: 1.2_
  - [x] 5.2 Add History toggle button to header

    - Add button to show/hide history panel
    - _Requirements: 4.1_
  - [x] 5.3 Implement toggleHistoryPanel function

    - Toggle panel visibility, update button state
    - _Requirements: 4.1, 4.2, 4.3_
  - [x] 5.4 Bind Undo button click handler

    - Call performUndo, update UI, show toast
    - _Requirements: 1.2_
  - [x] 5.5 Update Undo button state based on history

    - Disable when history is empty
    - _Requirements: 1.4_

- [ ] 6. Handle edge cases and errors
  - [x] 6.1 Add error handling for undo failures

    - Show error toast, retain entry in history
    - _Requirements: 1.5_
  - [x] 6.2 Ensure clearChat retains history

    - Update clearChat to not clear history
    - _Requirements: 3.2_

  - [ ] 6.3 Add empty state message to history panel
    - Show "No actions yet" when history is empty
    - _Requirements: 2.3_

- [x] 7. Checkpoint - Make sure all tests are passing


  - Ensure all tests pass, ask the user if questions arise.
