# Development Build Process

This document explains how the build pipeline works in this repository, from TypeScript compilation to VBA add-in artifacts.

## Build Commands at a Glance

- `npm run build` compiles the CLI/library into `lib/` and ensures vendor runtime assets.
- `npm run build:bootstrap` builds the bootstrap workbook (`scripts/bootstrap/build/bootstrap.xlsm`).
- `npm run build:addins` builds the production add-ins in `addins/build/` using the bootstrap workbook as the automation host.
- `npm version` chains `build` + `build:addins` + package creation.

The e2e scripts run `pnpm run build:check` first. It checks `lib/`,
`addins/build/vbapm.xlam` and `scripts/bootstrap/build/bootstrap.xlsm` in that
order, rebuilding only stale outputs. The add-in and bootstrap checks keep
SHA-256 freshness records beside their outputs. `pnpm run build:bootstrap`
remains available when a bootstrap refresh is needed explicitly.

## Bootstraping

The approach used in this project to build VBA documents is to have a VBA addin that lives within the VBA environment and acts as the relay between the package manager CLI and the VBA host.

However, in order to build that VBA addin in our build pipeline, we need a way to build a VBA document! This is the classic chicken-and-egg problem. We solve it by including the latest version of Bootstrap.xlsm (`scripts/bootstrap/build/bootstrap.xlsm`) inside the repo such that this file can be used to build the VBA addin with the latest source code.

Update flow for vbapm.xlam and bootstrap.xlsm
```mermaid
flowchart LR
    A@{ shape: doc, label: "Original Bootstrap.xlsm workbook" }
    X1[Original addin source code `/addins`]

    B@{ shape: doc, label: "Original vbapm.xlam addin" }

    A & X1 --> B1
    B1((build:addins))

    B1 --> B

    %%B & X1 --> B3
    %%B3((build:bootstrap))
    %%B3 --> A2



    A@{ shape: doc, label: "Original Bootstrap.xlsm workbook" }

    X[🔄️Addin source code update `/addins`]
    B2((build:addin))
    A & X  --> B2
    B2 --> BO@{ shape: doc, label: "Latest vbapm.xlam addin" }
    X & BO --> B4((build:bootstrap))
    AF@{ shape: doc, label: "Latest Bootstrap.xlsm workbook" }
    B4 --> AF

    %%AF --> A

```

## Windows Integration CI Workflow

In the Windows integration workflow, CI currently runs:

- `npm install`
- `npm run build`
- `npm run build:addins`

Because `scripts/bootstrap/build/bootstrap.xlsm` is already present in the repo, CI can build add-ins without first running `build:bootstrap`.

## Development Build Process (Practical Flow)

For day-to-day development, the expected flow is:

1. Build the JS/TS library (`npm run build`).
2. Use the committed bootstrap file to build the current add-in (`npm run build:addins`).
3. Optionally regenerate bootstrap (`npm run build:bootstrap`) when intentionally refreshing the seed host.

This gives both:

- Self-hosting evolution of the compiler/add-in code.
- A stable, committed starting point so new machines and CI are not blocked.
