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
});
