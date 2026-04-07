import { formatXml } from "../xml";

// Minimal XML fixture reproducing the entity-in-attribute-value bug found in
// OOXML table files (e.g. xl/tables/table1.xml). These files contain attribute
// values with &amp;, &lt;, &gt;, and the text-node string concatenation
// operator &amp; as produced by Excel's formula engine.
const TABLE_FIXTURE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table><autoFilter><filterColumn><filters><filter val="[$] 1.10 PEACE &amp; LOV"/><filter val="[€] 1.10 PEACE &amp; LOV"/></filters></filterColumn></autoFilter><tableColumns><tableColumn><calculatedColumnFormula>IFERROR(MATCH(Tbl[[#This Row],[Id]],IDs,0),0)&gt;0</calculatedColumnFormula></tableColumn><tableColumn><calculatedColumnFormula>LEFT(Tbl[[#This Row],[Name]],20) &amp; "…"</calculatedColumnFormula></tableColumn></tableColumns></table>`;

describe("formatXml – XML entity round-trip", () => {
	test("preserves &amp; in attribute values", () => {
		const result = formatXml(TABLE_FIXTURE);
		// Attribute values must NOT contain bare & (that would be invalid XML)
		// The filter val attributes contain & which must stay as &amp;
		expect(result).toContain(`val="[$] 1.10 PEACE &amp; LOV"`);
		expect(result).toContain(`val="[€] 1.10 PEACE &amp; LOV"`);
		expect(result).not.toMatch(/val="[^"]*[^;]&[^a][^"]*"/);
	});

	test("preserves &gt; in text content", () => {
		const result = formatXml(TABLE_FIXTURE);
		expect(result).toContain("&gt;0");
		expect(result).not.toContain(">0");
	});

	test("preserves &amp; in text content", () => {
		const result = formatXml(TABLE_FIXTURE);
		expect(result).toContain('&amp; "…"');
	});

	test("preserves &quot; in attribute values without double-escaping", () => {
		const xml = `<root val="A &quot; B"/>`;
		const result = formatXml(xml);
		expect(result).toBe(xml + "\r\n");
		expect(result).not.toContain("&amp;quot;");
	});

	test("preserves &lt; in attribute values", () => {
		const xml = `<root val="A &lt; B"/>`;
		const result = formatXml(xml);
		expect(result).toBe(xml + "\r\n");
	});

	test("preserves multiple entities in a single attribute value", () => {
		const xml = `<root val="A &amp; B &lt; C &gt; D"/>`;
		const result = formatXml(xml);
		expect(result).toBe(xml + "\r\n");
	});

	test("preserves whitespace-only text with xml:space=preserve", () => {
		const xml = `<si><t xml:space="preserve">   </t></si>`;
		const result = formatXml(xml);
		expect(result).toContain(`<t xml:space="preserve">   </t>`);
	});

	test("formatted output is valid (re-parseable) XML", () => {
		const result = formatXml(TABLE_FIXTURE);
		// Should not throw when re-parsed
		expect(() => formatXml(result)).not.toThrow();
	});
});
