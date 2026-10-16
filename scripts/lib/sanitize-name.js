/**
 * sanitize-name.js — Utility function to sanitize a string for use as a filename.
 *
 * This is used to generate the .block archive filename when publishing a package.
 *
 * It replaces "/" with "--" to allow for scoped package names (e.g. "@scope/name"),
 * and then uses the "sanitize-filename" library to ensure the resulting string is
 * safe for use as a filename across different operating systems.
 */

import sanitizeFilename from "sanitize-filename";

export default function sanitizeName(name) {
	return sanitizeFilename(name.replace("/", "--"), { replacement: "-" });
};
