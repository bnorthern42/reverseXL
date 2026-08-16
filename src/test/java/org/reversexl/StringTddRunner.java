package org.reversexl;

import java.util.Objects;

/**
 * Lightweight, zero-dependency string-based test runner that eliminates framework dependencies
 * like JUnit or TestNG to maintain a hermetic supply-chain posture.
 */
public class StringTddRunner {

    private static int passedCount = 0;
    private static int failedCount = 0;

    /**
     * Asserts that two objects are equal and prints formatted green/red test status to stdout.
     *
     * @param testName descriptive name of the test case
     * @param expected expected object value
     * @param actual actual evaluated value
     */
    public static void assertEquals(String testName, Object expected, Object actual) {
        if (Objects.equals(expected, actual)) {
            passedCount++;
            System.out.println("[GREEN] " + testName + " passed.");
        } else {
            failedCount++;
            System.out.println("[RED] " + testName + " FAILED: Expected '" + expected + "' but got '" + actual + "'.");
        }
    }

    /**
     * Returns the total count of passed test assertions in the current run.
     *
     * @return count of passed assertions
     */
    public static int getPassedCount() {
        return passedCount;
    }

    /**
     * Returns the total count of failed test assertions in the current run.
     *
     * @return count of failed assertions
     */
    public static int getFailedCount() {
        return failedCount;
    }

    /**
     * Resets the passed and failed counters to zero before running a new test suite.
     */
    public static void reset() {
        passedCount = 0;
        failedCount = 0;
    }

    /**
     * Prints a formatted summary block displaying the total passed and failed counts.
     */
    public static void summary() {
        System.out.println("\n-------------------------------------------------");
        System.out.println("TDD Test Results: " + passedCount + " passed, " + failedCount + " failed.");
        System.out.println("-------------------------------------------------");
    }
}
