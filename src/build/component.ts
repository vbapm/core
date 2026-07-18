import { CliError, ErrorCode } from "../errors";
import { readFile } from "../utils/fs";
import { extname } from "../utils/path";
import { BY_LINE } from "../utils/text";
import {
	Codepage,
	codepageToLabel,
	decodeBuffer,
	SniffResult,
	sniffEncoding
} from "./encoding-sniffer";
import * as iconv from "iconv-lite";

export type ComponentType = "module" | "class" | "form" | "object";

export interface ComponentDetails {
	path?: string;
	binary?: Buffer;
	/** Source encoding declared in TOML (for transcoding). */
	sourceEncoding?: string;
}

/**
 * Represents a VBA component, which can be a module, class, form, or document.
 *
 * The code is stored as a JavaScript string. When constructed from a Buffer,
 * the encoding is auto-detected via {@link sniffEncoding} so that files
 * exported by VBA in a locale-specific codepage (e.g. Windows-1252) are
 * decoded correctly.
 */
export class Component {
	type: ComponentType;
	code: string;
	details: ComponentDetails;
	/** Encoding that was detected when the component was loaded from a Buffer. */
	encoding?: SniffResult;

	constructor(
		type: ComponentType,
		code: Buffer | string,
		codepage: Codepage,
		details: ComponentDetails = {}
	) {
		this.type = type;

		if (code && Buffer.isBuffer(code)) {
			if (codepage === Codepage.Unknown) {
				const result = sniffEncoding(code);
				this.code = decodeBuffer(code, result);
				this.encoding = result;
			} else {
				this.code = iconv.decode(code, codepageToLabel(codepage));
			}
		} else {
			this.code = code as string;
		}

		this.details = details;
	}

	get name(): string {
		const line = findLine(this.code, "Attribute VB_Name");
		if (!line) {
			throw new CliError(
				ErrorCode.ComponentInvalidNoName,
				`Invalid component: No attribute VB_Name found.`
			);
		}

		const [, value] = line.split("=");
		return JSON.parse(value);
	}

	get binaryPath(): string | undefined {
		const line = findLine(this.code, "OleObjectBlob");
		if (!line) return;

		const [, value] = line.split("=", 2);
		const [path] = value.split(":", 2);
		return JSON.parse(path);
	}

	get filename(): string {
		const extension = typeToExtension[this.type];
		return `${this.name}${extension}`;
	}

	/**
	 * Load a component from a file on disk.
	 *
	 * @param path - Absolute path to the .bas / .cls / .frm file.
	 * @param codepage - Known codepage (e.g. from {@link getSystemCodepage}
	 *   when loading files just exported by VBA), or {@link Codepage.Unknown}
	 *   to auto-detect via sniffing.
	 * @param details - Optional binary path for .frx companions.
	 */
	static async load(
		path: string,
		codepage: Codepage,
		details: { binary_path?: string } = {}
	): Promise<Component> {
		const { binary_path } = details;

		const type = extensionToType[extname(path)];
		if (!type) {
			throw new CliError(
				ErrorCode.ComponentUnrecognized,
				`Unrecognized component extension "${extname(path)}" (at "${path}").`
			);
		}

		const code = await readFile(path);
		const binary = <Buffer | undefined>(binary_path && (await readFile(binary_path)));

		return new Component(type, code, codepage, { path, binary });
	}
}

export const extensionToType: { [extension: string]: ComponentType } = {
	".bas": "module",
	".cls": "class",
	".frm": "form"
};
export const typeToExtension: { [type: string]: string } = {
	module: ".bas",
	class: ".cls",
	form: ".frm",
	object: ".cls"
};

function findLine(code: string, search: string): string | undefined {
	const lines = code.split(BY_LINE).map(line => line.trim());
	return lines.find(line => line.startsWith(search));
}

export function byComponentName(a: Component, b: Component): number {
	if (a.name < b.name) return -1;
	if (a.name > b.name) return 1;
	return 0;
}

/** Order for sorting components by type: Objects → Modules → Forms → Classes. */
const TYPE_ORDER: Record<ComponentType, number> = {
	object: 1,
	module: 2,
	form: 3,
	class: 4
};

/**
 * Sort components by type (Objects → Modules → Forms → Classes),
 * then alphabetically by name within each type.
 */
export function byComponentTypeThenName(a: Component, b: Component): number {
	const orderA = TYPE_ORDER[a.type] ?? 99;
	const orderB = TYPE_ORDER[b.type] ?? 99;
	if (orderA !== orderB) return orderA - orderB;
	return byComponentName(a, b);
}
