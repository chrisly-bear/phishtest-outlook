/*
 * Unit tests for the header parser function.
 *
 * These tests verify the parseHeader function extracted from taskpane.js
 * works correctly for various RFC 5322 header scenarios.
 *
 * Run with: node test/parser.test.js
 */

// --- Extract the parseHeader function for testing ---
// We re-implement it here since taskpane.js depends on Office.js which
// isn't available in Node. The logic is identical.

function parseHeader(rawHeaders, headerName) {
    if (!rawHeaders || !headerName) {
        return null;
    }

    // Unfold RFC 5322 continuation lines
    const unfolded = rawHeaders.replace(/\r?\n[ \t]+/g, " ");

    // Split into individual header lines
    const lines = unfolded.split(/\r?\n/);

    const target = headerName.toLowerCase();
    const values = [];

    for (const line of lines) {
        const colonIndex = line.indexOf(":");
        if (colonIndex === -1) {
            continue;
        }

        const name = line.substring(0, colonIndex).trim().toLowerCase();
        const value = line.substring(colonIndex + 1).trim();

        if (name === target) {
            values.push(value);
        }
    }

    if (values.length === 0) {
        return null;
    }

    return values.join("; ");
}

// --- Test framework ---

let passed = 0;
let failed = 0;

function assert(condition, message) {
    if (condition) {
        console.log("  PASS: " + message);
        passed++;
    } else {
        console.error("  FAIL: " + message);
        failed++;
    }
}

function assertEqual(actual, expected, message) {
    const cond = actual === expected;
    if (cond) {
        console.log("  PASS: " + message);
        passed++;
    } else {
        console.error("  FAIL: " + message);
        console.error("    Expected: " + JSON.stringify(expected));
        console.error("    Actual:   " + JSON.stringify(actual));
        failed++;
    }
}

// --- Tests ---

console.log("\nTest 1: Simple X-PHISHTEST header present");
{
    const headers = [
        "From: sender@example.com",
        "To: recipient@example.com",
        "Subject: Test email",
        "X-PHISHTEST: This is a phishing test warning",
        "Date: Fri, 14 Aug 2026 10:00:00 +0200",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assertEqual(result, "This is a phishing test warning", "Should extract the header value");
}

console.log("\nTest 2: No X-PHISHTEST header");
{
    const headers = [
        "From: sender@example.com",
        "To: recipient@example.com",
        "Subject: Normal email",
        "Date: Fri, 14 Aug 2026 10:00:00 +0200",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assertEqual(result, null, "Should return null when header is absent");
}

console.log("\nTest 3: Case-insensitive header name (X-PhishTest)");
{
    const headers = [
        "From: sender@example.com",
        "X-PhishTest: Mixed case header",
        "Subject: Test",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assertEqual(result, "Mixed case header", "Should match case-insensitively");
}

console.log("\nTest 4: Case-insensitive lookup name (X-PHISHTEST in headers, x-phishtest lookup)");
{
    const headers = [
        "X-PHISHTEST: Uppercase in headers",
    ].join("\r\n");

    const result = parseHeader(headers, "X-PHISHTEST");
    assertEqual(result, "Uppercase in headers", "Should match regardless of lookup case");
}

console.log("\nTest 5: RFC 5322 header folding (continuation line)");
{
    const headers = [
        "From: sender@example.com",
        "X-PHISHTEST: This is a long warning that continues",
        "  on the next line with folding",
        "Subject: Test",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assertEqual(
        result,
        "This is a long warning that continues on the next line with folding",
        "Should unfold continuation lines"
    );
}

console.log("\nTest 6: Multiple X-PHISHTEST headers");
{
    const headers = [
        "X-PHISHTEST: First warning",
        "From: sender@example.com",
        "X-PHISHTEST: Second warning",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assertEqual(result, "First warning; Second warning", "Should join multiple values with '; '");
}

console.log("\nTest 7: Empty header value");
{
    const headers = [
        "X-PHISHTEST:",
        "From: sender@example.com",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assertEqual(result, "", "Should return empty string for empty value (not null)");
}

console.log("\nTest 7b: Empty header value should be !== null (detected as present)");
{
    const headers = [
        "X-PHISHTEST:",
        "From: sender@example.com",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assert(result !== null, "Empty header should be detected as present (not null)");
}

console.log("\nTest 8: Header name with extra whitespace around colon");
{
    const headers = [
        "X-PHISHTEST  :  Value with spaces around colon  ",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assertEqual(result, "Value with spaces around colon", "Should trim whitespace from name and value");
}

console.log("\nTest 9: Empty raw headers string");
{
    const result = parseHeader("", "x-phishtest");
    assertEqual(result, null, "Should return null for empty input");
}

console.log("\nTest 10: Null raw headers");
{
    const result = parseHeader(null, "x-phishtest");
    assertEqual(result, null, "Should return null for null input");
}

console.log("\nTest 11: Header with tab continuation (folding with tab)");
{
    const headers = [
        "X-PHISHTEST: First part of warning",
        "\tsecond part after tab",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assertEqual(
        result,
        "First part of warning second part after tab",
        "Should unfold tab-continuation lines"
    );
}

console.log("\nTest 12: Header name that is a prefix of another header");
{
    const headers = [
        "X-PHISHTEST: Correct header",
        "X-PHISHTESTING: Should not match",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assertEqual(result, "Correct header", "Should match exact header name only");
}

console.log("\nTest 13: \n in header value (should not break parsing)");
{
    const headers = [
        "X-PHISHTEST: Warning with \\n literal characters",
    ].join("\r\n");

    const result = parseHeader(headers, "x-phishtest");
    assertEqual(result, "Warning with \\n literal characters", "Should handle backslash-n in value");
}

console.log("\nTest 14: Long header value (over 150 chars, truncation handled by caller)");
{
    const longValue = "A".repeat(200);
    const headers = "X-PHISHTEST: " + longValue;

    const result = parseHeader(headers, "x-phishtest");
    assert(result !== null, "Should find the header even if value is very long");
    assertEqual(result.length, 200, "Should return full value (truncation is caller's responsibility)");
}

// --- Summary ---
console.log("\n" + "=".repeat(50));
console.log("Results: " + passed + " passed, " + failed + " failed");
console.log("=".repeat(50));

process.exit(failed > 0 ? 1 : 0);
