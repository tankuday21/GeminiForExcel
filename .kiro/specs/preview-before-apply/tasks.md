# Implementation Plan

- [x] 1. Set up preview panel foundation



  - [x] 1.1 Create preview panel HTML structure in taskpane.html

    - Add preview panel container with collapsible action list
    - Add checkbox inputs for each action item
    - _Requirements: 1.1, 2.1, 3.2_

  - [x] 1.2 Add preview panel CSS styles in taskpane.css

    - Style preview items, checkboxes, icons, expand/collapse states
    - _Requirements: 3.1, 3.2, 3.3, 3.4_

  - [x] 1.3 Implement getActionIcon function

    - Return distinct SVG icons for formula, values, format, chart, validation types
    - _Requirements: 3.1_
  - [x] 1.4 Write property test for action icons


    - **Property 5: Action type maps to distinct icon**
    - **Validates: Requirements 3.1**


- [ ] 2. Implement preview rendering logic
  - [x] 2.1 Implement getActionSummary function

    - Return human-readable summary with action type and target
    - _Requirements: 3.4_

  - [ ] 2.2 Implement getActionDetails function
    - Return full details including all action properties

    - _Requirements: 3.3_
  - [ ] 2.3 Implement renderPreviewItem function
    - Render single action with checkbox, icon, summary/details based on expanded state

    - _Requirements: 1.2, 1.3, 1.4, 1.5, 1.6, 3.3, 3.4_
  - [x] 2.4 Write property test for action rendering

    - **Property 2: Action rendering includes required fields**
    - **Validates: Requirements 1.2, 1.3, 1.4, 1.5, 1.6**

  - [ ] 2.5 Write property test for collapsed view
    - **Property 6: Collapsed view shows summary**
    - **Validates: Requirements 3.4**
  - [ ] 2.6 Write property test for expanded view
    - **Property 7: Expanded view shows full details**
    - **Validates: Requirements 3.3**

- [-] 3. Implement preview panel component


  - [ ] 3.1 Implement renderPreviewPanel function
    - Render all actions with checkboxes, handle expand/collapse
    - _Requirements: 1.1, 2.1, 3.2_

  - [ ] 3.2 Write property test for preview panel
    - **Property 1: Preview renders all actions**

    - **Validates: Requirements 1.1**
  - [ ] 3.3 Write property test for checkboxes
    - **Property 3: Each action has a checkbox**
    - **Validates: Requirements 2.1**

- [-] 4. Implement selection management


  - [ ] 4.1 Implement filterSelectedActions function
    - Filter actions array based on selection boolean array
    - _Requirements: 2.2, 2.3_

  - [ ] 4.2 Implement hasSelectedActions function
    - Return true if any selection is true
    - _Requirements: 2.4_
  - [x] 4.3 Add PreviewState to application state


    - Track selections, expandedIndex, highlightedIndex

    - _Requirements: 2.1, 2.2_
  - [ ] 4.4 Write property test for filter function
    - **Property 4: Filter returns only selected actions**
    - **Validates: Requirements 2.2, 2.3**



- [ ] 5. Integrate preview panel with existing flow
  - [x] 5.1 Update handleSend to show preview panel when actions exist

    - Show preview panel after AI response with actions
    - _Requirements: 1.1_
  - [x] 5.2 Update handleApply to use selected actions only

    - Apply only checked actions from preview
    - _Requirements: 2.3_

  - [ ] 5.3 Update Apply button state based on selections
    - Disable when no actions selected

    - _Requirements: 2.4_
  - [ ] 5.4 Add expand/collapse toggle functionality
    - Click to expand/collapse action details
    - _Requirements: 3.2, 3.3, 3.4_



- [ ] 6. Implement Excel highlighting
  - [ ] 6.1 Implement highlightRange function
    - Use Excel API to select/highlight target range

    - _Requirements: 4.1_
  - [ ] 6.2 Implement clearHighlight function
    - Clear any active highlighting
    - _Requirements: 4.2_

  - [ ] 6.3 Add hover event handlers to preview items
    - Highlight on mouseenter, clear on mouseleave
    - _Requirements: 4.1, 4.2_
  - [x] 6.4 Add error handling for invalid ranges

    - Show warning indicator if highlight fails
    - _Requirements: 4.3_


- [x] 7. Checkpoint - Make sure all tests are passing

  - Ensure all tests pass, ask the user if questions arise.
