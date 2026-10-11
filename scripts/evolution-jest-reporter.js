const { existsSync, readFileSync, writeFileSync } = require("node:fs");
const { join } = require("node:path");

class EvolutionJestReporter {
	onRunComplete(_contexts, results) {
		try {
			const marker = join(
				process.env.TEMP || process.env.TMP || ".",
				"Excel-Instances",
				"evolution",
				"current-run.txt"
			);
			if (!existsSync(marker)) return;
			const runDir = readFileSync(marker, "utf8").trim();
			const infoPath = join(runDir, "run-info.json");
			if (!runDir || !existsSync(infoPath)) return;

			const failures = [];
			for (const testFile of results.testResults || []) {
				for (const assertion of testFile.testResults || []) {
					if (assertion.status === "failed") {
						failures.push({
							testFilePath: testFile.testFilePath,
							fullName: assertion.fullName,
							failureMessages: assertion.failureMessages || []
						});
					}
				}
			}

			const info = JSON.parse(readFileSync(infoPath, "utf8"));
			const snapshot = results.snapshot || {};
			// NOTE: results.success is still `false` when reporters' onRunComplete
			// fires - Jest assigns the real value only *after* dispatching
			// onRunComplete (TestScheduler in @jest/core). Derive success from the
			// counters and snapshot.failure (which Jest sets before the dispatch)
			// so it mirrors the actual exit status.
			const success =
				results.numFailedTests === 0 &&
				results.numFailedTestSuites === 0 &&
				!snapshot.failure &&
				!results.runExecError;
			info.jest = {
				success,
				numTotalTestSuites: results.numTotalTestSuites,
				numPassedTestSuites: results.numPassedTestSuites,
				numFailedTestSuites: results.numFailedTestSuites,
				numPendingTestSuites: results.numPendingTestSuites,
				numTotalTests: results.numTotalTests,
				numPassedTests: results.numPassedTests,
				numFailedTests: results.numFailedTests,
				numPendingTests: results.numPendingTests,
				numTodoTests: results.numTodoTests,
				numTotalSnapshots: snapshot.total || 0,
				numAddedSnapshots: snapshot.added || 0,
				numUpdatedSnapshots: snapshot.updated || 0,
				numUnmatchedSnapshots: snapshot.unmatched || 0,
				runExecError: results.runExecError || "",
				failures
			};
			writeFileSync(infoPath, JSON.stringify(info, null, 2), "utf8");
		} catch (error) {
			process.stderr.write(`[evolution-jest-reporter] ${error.message}\n`);
		}
	}
}

module.exports = EvolutionJestReporter;
