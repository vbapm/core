import { Codepage } from "../encoding-sniffer";
import { Component } from "../component";

// ── helpers ──────────────────────────────────────────────────────

function cp1252(text: string): Buffer {
	const MAP: Record<string, number> = {
		"é": 0xe9, "è": 0xe8, "ê": 0xea, "ë": 0xeb,
		"à": 0xe0, "â": 0xe2, "ä": 0xe4,
		"ù": 0xf9, "û": 0xfb, "ü": 0xfc,
		"ç": 0xe7, "ô": 0xf4, "î": 0xee, "ï": 0xef,
		"É": 0xc9, "È": 0xc8, "Ê": 0xca, "Ë": 0xcb,
		"À": 0xc0, "Â": 0xc2, "Ä": 0xc4,
		"Ù": 0xd9, "Û": 0xdb, "Ü": 0xdc,
		"Ç": 0xc7, "œ": 0x9c, "Œ": 0x8c, "€": 0x80,
	};

	const bytes: number[] = [];
	for (let i = 0; i < text.length; i++) {
		const ch = text[i];
		const cp = MAP[ch];
		bytes.push(cp !== undefined ? cp : ch.charCodeAt(0));
	}
	return Buffer.from(bytes);
}

function makeVbaModule(name: string, body: string): string {
	return [
		`Attribute VB_Name = "${name}"`,
		"",
		body
	].join("\n");
}

// ── tests ────────────────────────────────────────────────────────

describe("Component encoding", () => {
	const vbaBody = "' Démonstration\nPublic Sub Test()\n    MsgBox \"Voilà !\"\nEnd Sub\n";

	test("CP1252 buffer + Codepage.Windows1252 → correct accents", () => {
		const source = makeVbaModule("ModuleAccents", vbaBody);
		const buf = cp1252(source);
		const comp = new Component("module", buf, Codepage.Windows1252);

		expect(comp.name).toBe("ModuleAccents");
		expect(comp.code).toBe(source);
		expect(comp.code).not.toContain("\uFFFD");
		expect(comp.encoding).toBeUndefined();
	});

	test("CP1252 buffer + Codepage.Unknown → correct accents (sniffed)", () => {
		const source = makeVbaModule("ModuleAccents", vbaBody);
		const buf = cp1252(source);
		const comp = new Component("module", buf, Codepage.Unknown);

		expect(comp.name).toBe("ModuleAccents");
		expect(comp.code).toBe(source);
		expect(comp.code).not.toContain("\uFFFD");
		expect(comp.encoding).toBeDefined();
		expect(comp.encoding!.encoding).toBe("windows-1252");
	});

	test("UTF-8 buffer + Codepage.Unknown → correct", () => {
		const source = makeVbaModule("ModuleTest", "' Hello World");
		const buf = Buffer.from(source, "utf8");
		const comp = new Component("module", buf, Codepage.Unknown);

		expect(comp.name).toBe("ModuleTest");
		expect(comp.code).toBe(source);
		expect(comp.encoding!.encoding).toBe("utf8");
	});

	test("UTF-8 buffer with French + Codepage.Unknown → correct", () => {
		const source = makeVbaModule("Modèle", "' Résumé");
		const buf = Buffer.from(source, "utf8");
		const comp = new Component("module", buf, Codepage.Unknown);

		expect(comp.name).toBe("Modèle");
		expect(comp.code).toBe(source);
		expect(comp.encoding!.encoding).toBe("utf8");
	});

	test("string code (not Buffer) → no encoding applied", () => {
		const source = makeVbaModule("ModuleTest", "' Hello");
		const comp = new Component("module", source, Codepage.Unknown);

		expect(comp.code).toBe(source);
		expect(comp.encoding).toBeUndefined();
	});

	test("class component with CP1252 accents + known codepage", () => {
		const source = [
			"VERSION 1.0 CLASS",
			"Attribute VB_Name = \"ClasseRésumé\"",
			"",
			"Public Property Get Résultat() As String",
			"    Résultat = \"Opération réussie\"",
			"End Property"
		].join("\n");
		const buf = cp1252(source);
		const comp = new Component("class", buf, Codepage.Windows1252);

		expect(comp.name).toBe("ClasseRésumé");
		expect(comp.code).toBe(source);
		expect(comp.code).not.toContain("\uFFFD");
	});

	test("form component with CP1252 + Unknown codepage → correct .frm", () => {
		const source = [
			"VERSION 5.00",
			"Attribute VB_Name = \"UserForm1\"",
			"",
			"' Formulaire de démarrage"
		].join("\n");
		const buf = cp1252(source);
		const comp = new Component("form", buf, Codepage.Unknown);

		expect(comp.name).toBe("UserForm1");
		expect(comp.code).toBe(source);
	});

	test("CP1252 edge chars: œ Œ € → decoded correctly with known codepage", () => {
		const source = makeVbaModule("Test", "' Cœur €100");
		const buf = cp1252(source);
		const comp = new Component("module", buf, Codepage.Windows1252);

		expect(comp.code).toBe(source);
		expect(comp.code).not.toContain("\uFFFD");
	});
});
