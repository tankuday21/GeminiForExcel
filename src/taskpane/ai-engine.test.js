/**
 * Tests for AI Engine - Advanced AI capabilities
 */

import {
    detectTaskType,
    TASK_TYPES,
    isCorrection,
    searchPatterns,
    requiresMultiStep,
    decomposeTask,
    parseFunctionCalls
} from "./ai-engine.js";

describe("AI Engine", () => {
    describe("detectTaskType", () => {
        test("detects formula tasks", () => {
            expect(detectTaskType("Create a SUM formula")).toBe(TASK_TYPES.FORMULA);
            expect(detectTaskType("Calculate the average")).toBe(TASK_TYPES.FORMULA);
            expect(detectTaskType("Add up all values")).toBe(TASK_TYPES.FORMULA);
            expect(detectTaskType("Use VLOOKUP to find")).toBe(TASK_TYPES.FORMULA);
        });

        test("detects chart tasks", () => {
            expect(detectTaskType("Create a bar chart")).toBe(TASK_TYPES.CHART);
            expect(detectTaskType("Visualize this data")).toBe(TASK_TYPES.CHART);
            expect(detectTaskType("Make a pie graph")).toBe(TASK_TYPES.CHART);
        });

        test("detects analysis tasks", () => {
            expect(detectTaskType("Analyze this data")).toBe(TASK_TYPES.ANALYSIS);
            expect(detectTaskType("Give me insights")).toBe(TASK_TYPES.ANALYSIS);
            expect(detectTaskType("Find patterns in the data")).toBe(TASK_TYPES.ANALYSIS);
        });

        test("detects format tasks", () => {
            expect(detectTaskType("Format as bold")).toBe(TASK_TYPES.FORMAT);
            expect(detectTaskType("Change the color")).toBe(TASK_TYPES.FORMAT);
            expect(detectTaskType("Style the header")).toBe(TASK_TYPES.FORMAT);
        });

        test("detects validation tasks", () => {
            expect(detectTaskType("Create a dropdown")).toBe(TASK_TYPES.VALIDATION);
            expect(detectTaskType("Add validation")).toBe(TASK_TYPES.VALIDATION);
        });

        test("detects table tasks", () => {
            expect(detectTaskType("Create a table")).toBe(TASK_TYPES.TABLE);
            expect(detectTaskType("Format as table")).toBe(TASK_TYPES.TABLE);
            expect(detectTaskType("Add table column")).toBe(TASK_TYPES.TABLE);
            expect(detectTaskType("Convert to table")).toBe(TASK_TYPES.TABLE);
        });

        test("detects pivot tasks", () => {
            expect(detectTaskType("Create pivot table")).toBe(TASK_TYPES.PIVOT);
            expect(detectTaskType("Add pivot field")).toBe(TASK_TYPES.PIVOT);
            expect(detectTaskType("Summarize with pivot")).toBe(TASK_TYPES.PIVOT);
        });

        test("detects data manipulation tasks", () => {
            expect(detectTaskType("Insert rows at row 5")).toBe(TASK_TYPES.DATA_MANIPULATION);
            expect(detectTaskType("Merge cells A1 to C1")).toBe(TASK_TYPES.DATA_MANIPULATION);
            expect(detectTaskType("Find and replace")).toBe(TASK_TYPES.DATA_MANIPULATION);
            expect(detectTaskType("Text to columns")).toBe(TASK_TYPES.DATA_MANIPULATION);
        });

        test("detects shape tasks", () => {
            expect(detectTaskType("Insert rectangle")).toBe(TASK_TYPES.SHAPES);
            expect(detectTaskType("Add image")).toBe(TASK_TYPES.SHAPES);
            expect(detectTaskType("Create text box")).toBe(TASK_TYPES.SHAPES);
        });

        test("detects comment tasks", () => {
            expect(detectTaskType("Add comment")).toBe(TASK_TYPES.COMMENTS);
            expect(detectTaskType("Reply to comment")).toBe(TASK_TYPES.COMMENTS);
            expect(detectTaskType("Resolve comment")).toBe(TASK_TYPES.COMMENTS);
        });

        test("detects protection tasks", () => {
            expect(detectTaskType("Protect worksheet")).toBe(TASK_TYPES.PROTECTION);
            expect(detectTaskType("Lock cells")).toBe(TASK_TYPES.PROTECTION);
            expect(detectTaskType("Unprotect sheet")).toBe(TASK_TYPES.PROTECTION);
            expect(detectTaskType("Protect table")).toBe(TASK_TYPES.PROTECTION);
            expect(detectTaskType("Protect this data")).toBe(TASK_TYPES.PROTECTION);
        });

        test("detects page setup tasks", () => {
            expect(detectTaskType("Set page orientation")).toBe(TASK_TYPES.PAGE_SETUP);
            expect(detectTaskType("Add header and footer")).toBe(TASK_TYPES.PAGE_SETUP);
            expect(detectTaskType("Set print area")).toBe(TASK_TYPES.PAGE_SETUP);
        });

        test("handles priority correctly for overlapping keywords", () => {
            expect(detectTaskType("Create pivot table")).toBe(TASK_TYPES.PIVOT); // not TABLE
            expect(detectTaskType("Insert row in table")).toBe(TASK_TYPES.DATA_MANIPULATION); // not TABLE
            expect(detectTaskType("Protect sheet with password")).toBe(TASK_TYPES.PROTECTION); // not FORMAT
        });

        test("returns general for unknown tasks", () => {
            expect(detectTaskType("Hello")).toBe(TASK_TYPES.GENERAL);
            expect(detectTaskType("What can you do?")).toBe(TASK_TYPES.GENERAL);
        });
    });

    describe("isCorrection", () => {
        test("detects correction messages", () => {
            expect(isCorrection("No, column E not C")).toBe(true);
            expect(isCorrection("Wrong, use column B")).toBe(true);
            expect(isCorrection("Actually, I meant row 5")).toBe(true);
            expect(isCorrection("That's not right")).toBe(true);
            expect(isCorrection("Use column E instead")).toBe(true);
        });

        test("does not flag normal messages as corrections", () => {
            expect(isCorrection("Create a sum formula")).toBe(false);
            expect(isCorrection("Show me a chart")).toBe(false);
            expect(isCorrection("What is the total?")).toBe(false);
        });
    });

    describe("searchPatterns", () => {
        test("finds relevant patterns for sum queries", () => {
            const patterns = searchPatterns("sum all values");
            expect(patterns.length).toBeGreaterThan(0);
            expect(patterns[0].id).toBe("sum_column");
        });

        test("finds relevant patterns for lookup queries", () => {
            const patterns = searchPatterns("lookup a value");
            expect(patterns.length).toBeGreaterThan(0);
            expect(patterns.some(p => p.id.includes("lookup"))).toBe(true);
        });

        test("finds relevant patterns for percentage queries", () => {
            const patterns = searchPatterns("calculate percentage");
            expect(patterns.length).toBeGreaterThan(0);
            expect(patterns.some(p => p.id.includes("percentage"))).toBe(true);
        });

        test("limits results for unrelated queries", () => {
            const patterns = searchPatterns("xyz123abc");
            // May return some patterns due to partial word matching
            expect(patterns.length).toBeLessThanOrEqual(5);
        });
    });

    describe("requiresMultiStep", () => {
        test("identifies complex tasks", () => {
            expect(requiresMultiStep("Analyze and then create a chart")).toBe(true);
            expect(requiresMultiStep("Do multiple calculations")).toBe(true);
            expect(requiresMultiStep("Process all columns")).toBe(true);
        });

        test("identifies simple tasks", () => {
            expect(requiresMultiStep("Sum column A")).toBe(false);
            expect(requiresMultiStep("Create chart")).toBe(false);
        });

        test("identifies long prompts as complex", () => {
            const longPrompt = "a".repeat(250);
            expect(requiresMultiStep(longPrompt)).toBe(true);
        });
    });

    describe("decomposeTask", () => {
        test("creates steps for formula tasks", () => {
            const steps = decomposeTask("Create a complex formula", {});
            expect(steps.length).toBeGreaterThanOrEqual(3);
            expect(steps[0].step).toBe("analyze");
            expect(steps.some(s => s.step === "plan")).toBe(true);
            expect(steps.some(s => s.step === "execute")).toBe(true);
        });

        test("creates steps for chart tasks", () => {
            const steps = decomposeTask("Create a chart to visualize", {});
            expect(steps.length).toBeGreaterThanOrEqual(3);
        });

        test("creates steps for analysis tasks", () => {
            const steps = decomposeTask("Analyze this data", {});
            expect(steps.length).toBeGreaterThanOrEqual(3);
        });
    });

    describe("parseFunctionCalls", () => {
        test("parses function call syntax", () => {
            const response = 'CALL_FUNCTION("SUM", "A10", "A1:A9")';
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(1);
            expect(calls[0].type).toBe("formula");
            expect(calls[0].target).toBe("A10");
            expect(calls[0].data).toContain("=SUM");
        });

        test("parses multiple function calls", () => {
            const response = `
                CALL_FUNCTION("SUM", "A10", "A1:A9")
                CALL_FUNCTION("AVERAGE", "B10", "B1:B9")
            `;
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(2);
        });

        test("ignores invalid function names", () => {
            const response = 'CALL_FUNCTION("INVALID_FUNC", "A1", "test")';
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(0);
        });
    });

    describe("parseFunctionCalls - Excel 365 Functions", () => {
        test("parses FILTER function call", () => {
            const response = 'CALL_FUNCTION("FILTER", "E2", "A2:C100, B2:B100=\"Sales\", \"No results\"")';
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(1);
            expect(calls[0].type).toBe("formula");
            expect(calls[0].target).toBe("E2");
            expect(calls[0].data).toContain("=FILTER");
        });

        test("parses SORT function call", () => {
            const response = 'CALL_FUNCTION("SORT", "E2", "A2:C100, 2, -1")';
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(1);
            expect(calls[0].data).toContain("=SORT");
        });

        test("parses UNIQUE function call", () => {
            const response = 'CALL_FUNCTION("UNIQUE", "E2", "A2:A100")';
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(1);
            expect(calls[0].data).toContain("=UNIQUE");
        });

        test("parses XMATCH function call", () => {
            const response = 'CALL_FUNCTION("XMATCH", "D2", "\"Apple\", A:A, 0")';
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(1);
            expect(calls[0].data).toContain("=XMATCH");
        });

        test("parses TEXTSPLIT function call", () => {
            const response = 'CALL_FUNCTION("TEXTSPLIT", "B2", "A2, \",\"")';
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(1);
            expect(calls[0].data).toContain("=TEXTSPLIT");
        });

        test("parses CHOOSECOLS function call", () => {
            const response = 'CALL_FUNCTION("CHOOSECOLS", "F2", "A1:E100, 1, 3, 5")';
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(1);
            expect(calls[0].data).toContain("=CHOOSECOLS");
        });

        test("parses SEQUENCE function call", () => {
            const response = 'CALL_FUNCTION("SEQUENCE", "A1", "10, 1, 1, 1")';
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(1);
            expect(calls[0].data).toContain("=SEQUENCE");
        });

        test("parses TAKE function call", () => {
            const response = 'CALL_FUNCTION("TAKE", "G2", "A1:C100, 10")';
            const calls = parseFunctionCalls(response);
            expect(calls.length).toBe(1);
            expect(calls[0].data).toContain("=TAKE");
        });
    });

    describe("searchPatterns - Dynamic Array Patterns", () => {
        test("finds FILTER pattern for filter queries", () => {
            const patterns = searchPatterns("filter sales data");
            expect(patterns.some(p => p.id === "filter_by_criteria")).toBe(true);
        });

        test("finds SORT pattern for sort queries", () => {
            const patterns = searchPatterns("sort by amount dynamically");
            expect(patterns.some(p => p.id === "sort_dynamic")).toBe(true);
        });

        test("finds UNIQUE pattern for unique queries", () => {
            const patterns = searchPatterns("get unique values");
            expect(patterns.some(p => p.id === "unique_list")).toBe(true);
        });

        test("finds TEXTSPLIT pattern for split queries", () => {
            const patterns = searchPatterns("split text by comma");
            expect(patterns.some(p => p.id === "textsplit_parse")).toBe(true);
        });

        test("finds combo pattern for complex queries", () => {
            const patterns = searchPatterns("filter and sort data");
            expect(patterns.some(p => p.id === "filter_sort_combo")).toBe(true);
        });

        test("finds TAKE pattern for top queries", () => {
            const patterns = searchPatterns("get top 10 rows");
            expect(patterns.some(p => p.id === "take_top")).toBe(true);
        });

        test("finds GROUPBY pattern for group queries", () => {
            const patterns = searchPatterns("group sales by region");
            expect(patterns.some(p => p.id === "groupby_aggregate")).toBe(true);
        });

        test("finds PIVOTBY pattern for pivot queries", () => {
            const patterns = searchPatterns("create pivot summary");
            expect(patterns.some(p => p.id === "pivotby_summary")).toBe(true);
        });

        test("finds RANDARRAY pattern for random queries", () => {
            const patterns = searchPatterns("generate random numbers");
            expect(patterns.some(p => p.id === "randarray_generate")).toBe(true);
        });
    });

    describe("detectTaskType - Dynamic Array Keywords", () => {
        test("detects filter as formula task", () => {
            expect(detectTaskType("filter the data by sales")).toBe(TASK_TYPES.FORMULA);
        });

        test("detects unique as formula task", () => {
            expect(detectTaskType("get unique values from column")).toBe(TASK_TYPES.FORMULA);
        });

        test("detects textsplit as formula task", () => {
            expect(detectTaskType("split text by delimiter")).toBe(TASK_TYPES.FORMULA);
        });

        test("detects dynamic array as formula task", () => {
            expect(detectTaskType("use dynamic array formula")).toBe(TASK_TYPES.FORMULA);
        });
    });
});
