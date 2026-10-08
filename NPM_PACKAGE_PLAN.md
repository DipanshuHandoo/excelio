# ExcelIO npm Package Plan

## Decisions

- Package: `@dipanshuhandoo/excelio`.
- Initial version: `0.1.0`; MIT license; author: DipanshuHandoo.
- Support Node.js 22 and 24 LTS, ES modules, CommonJS, and TypeScript.
- Publish readable builds by default and expose minified builds through `/min`.
- Use local npm publishing commands and CI validation, not automatic publishing.
- Follow `worker_pool` conventions: esbuild, conditional exports, build/package
  verification scripts, public publish configuration, and lifecycle release gates.
- Confirm the GitHub repository URL, npm scope ownership, and package availability
  before release. Never publish or commit credentials as part of implementation.

## 1. Repository Layout

- Keep runtime sources, worker entrypoints, and existing tests under `src/`.
- Move README, LICENSE, and `.gitignore` to the repository root.
- Add root package metadata and lockfile, build/verification scripts, consumer
  fixtures, support documentation, changelog, and CI validation.
- Keep generated `dist/` files out of git; include them in the npm tarball.

## 2. Worker Compatibility

- Restore worker-generated `Uint8Array` values to `Buffer`, including arrays of
  workbook buffers, without changing file-target results.
- Keep readable inputs inline even when worker execution is explicitly requested.
- Preserve inline routing for hooks, signals, loggers, and writable destinations.
- Resolve workers relative to the installed entrypoint in both ESM and CommonJS.
- Reject unexpected worker exits; preserve structured errors across threads.
- Add regression tests before making release claims. A fresh worker per operation
  is worker offloading, not a reusable pool or parallel workbook chunking.

## 3. Build and Distribution

- Use esbuild as a development dependency; externalize `exceljs` and Node built-ins.
- Generate readable ESM/CommonJS and minified ESM/CommonJS output trees, each with
  its own two bundled worker entrypoints, independent of unpublished sources.
- Use `.js` for ESM and `.cjs` for CommonJS; explicitly support module-relative
  worker resolution in each format.
- Produce source maps and ESM/CommonJS declarations. Do not add browser bundles:
  filesystem, streams, and worker_threads make this a Node.js package.
- Configure root and `/min` conditional exports with matching `types`, `import`,
  and `require` targets, plus compatible `main`, `module`, and `types` fallbacks.
- Allowlist `dist`, README, LICENSE, and changelog through `files`.

## 4. Types and Dependencies

- Keep `exceljs` as the sole direct runtime dependency with a tested version range.
- Add esbuild, TypeScript, and Node type definitions as development dependencies;
  commit the generated npm lockfile.
- Import Buffer types explicitly; describe parsed workbook metadata separately
  from write specifications; declare the custom error constructor.
- Validate ESM and CommonJS consumers using TypeScript NodeNext resolution.

## 5. npm Commands

- `build`: clean and generate all distribution outputs and declarations.
- `build:min`: regenerate minified variants without removing readable outputs.
- `test`: existing smoke tests plus focused worker/stream regressions.
- `test:types`: compile ESM and CommonJS TypeScript consumers.
- `test:package`: install a real packed tarball in isolated consumer projects and
  verify ESM/CommonJS, minified exports, types, and worker execution.
- `test:perf`: optional existing 50,000-row benchmark, separate from release gates.
- `verify`: tests, clean build, types, and installed-package verification.
- `pack:check`: inspect and validate package contents.
- `publish:dry` / `publish:dry-run`: public npm dry run.
- `publish:public` / `release`: explicit public npm publication.
- `prepack`: build and validate artifacts.
- `prepublishOnly`: full verification. Internal packing uses `--ignore-scripts`
  to prevent recursive lifecycle execution.

## 6. Tests and CI

- Cover forced workers, configurable automatic thresholds, Buffer identity,
  multiple workbooks, inline stream/hook fallback, and structured worker errors.
- Cover file/stream writes, stream reads, transformations, row validation, date
  conversion, custom headers, formatting, progress, and cancellation checks.
- Pack the package and install it outside the repository so local sources and
  dependencies cannot mask missing distribution files.
- Exercise all four exports, declarations, and actual worker entrypoints.
- Check required files and exclude secrets, temporary spreadsheets, tests, and
  accidental dependencies on source-only files.
- Run CI on Node 22/24 and Linux/Windows using `npm ci` and `npm run verify`.
- Never store npm publishing credentials in CI validation or repository files.

## 7. Documentation and Support

- Document installation, import/require examples, minified exports, API contracts,
  workers, thresholds, hooks, cancellation, memory behavior, and support policy.
- Correct constant-memory claims: reading loads the workbook; input arrays,
  shared strings, and buffered output also consume memory.
- Document exclusions: browsers, legacy XLS, worker pooling/chunking, charts,
  macros, encryption, and guaranteed speed improvements.
- Add CONTRIBUTING, SECURITY, CHANGELOG, issue templates, and a release checklist.
- Use SemVer; document the pre-1.0 compatibility policy and supported runtimes.

## 8. Release Checklist

1. Confirm GitHub repository metadata, npm scope ownership, package availability,
   npm authentication, and account 2FA.
2. Review changes and select the version explicitly; do not automatically commit,
   tag, or publish as part of implementation.
3. Run `npm ci`, `npm run verify`, and `npm run publish:dry-run` from a clean checkout.
4. Inspect tarball contents and resolve every required validation failure.
5. Run `npm run release` only with explicit release approval.
6. Verify installation from the public registry and worker execution; record the
   matching GitHub release and tag. npm versions cannot be overwritten.

## Acceptance Criteria

- All distribution formats build reproducibly from the lockfile.
- Installed consumers can import/require root and minified exports and use workers.
- Worker writes preserve Buffer results and stream requests remain inline.
- Declarations compile for ESM/CommonJS; package contents are allowlisted.
- Required tests and CI pass; dry-run publishing contains all necessary files.
- Documentation describes implemented behavior, not unsupported pooling or
  constant-memory guarantees. No real npm publication occurs during setup.