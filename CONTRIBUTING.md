# Contributing

Thanks for contributing to vbapm.

## Development Setup

1. Run `npm install`
2. Run `npm run format`
3. Run `npm run build:cli`
4. Run `npm run build:addins`

## Testing

1. Run unit tests with `npm test`
2. Run end-to-end tests with `npm run test:e2e` or `npm run test:e2e:background`
3. To keep temporary e2e folders for inspection, set `KEEP_E2E_TMP=1` before running tests

## Verbose e2e Logs

e2e command output can be echoed even when tests succeed.

Use one of these options:

- Pass Jest verbose through npm args:
  - PowerShell: `npm run test:e2e:background -- --verbose`
- Set the environment variable used by the e2e helper:
  - PowerShell: `$env:E2E_VERBOSE=1; npm run test:e2e:background`

This prints each invoked e2e command plus its stdout and stderr, which helps compare local runs with CI logs.
