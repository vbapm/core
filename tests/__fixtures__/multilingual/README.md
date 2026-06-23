# Multilingual Encoding Fixtures

These fixtures are used by `tests/excel.mlang.ts` to verify that VBA source
files encoded in different Windows ANSI codepages survive the full build →
export roundtrip.

## Fixtures

| Directory | Codepage | Region |
|-----------|----------|--------|
| `cp1252/` | 1252 | Western European (French, German, Spanish) |
| `cp1251/` | 1251 | Cyrillic (Russian) |
| `cp1250/` | 1250 | Central European (Polish, Czech) |
| `cp932/`  | 932  | Japanese (CP932) |
| `cp936/`  | 936  | Simplified Chinese (GBK) |

## How it works

1. Each fixture is a minimal vbapm project with one `.bas` module encoded in
   the target codepage.
2. In CI, a Windows image is configured with the matching system locale
   (e.g. `fr-FR` for CP1252).
3. The test reads the system ANSI codepage from the registry and runs only
   the matching fixture.
4. The fixture is built (VBA Import), then exported (VBA Export), and the
   resulting `.bas` file is verified to contain the expected characters.

## Running

The test is gated behind the `E2E_ML` environment variable:

```powershell
pnpm test:e2e:mlang
```
