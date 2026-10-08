# Security Policy

## Reporting

Do not disclose vulnerabilities or sensitive workbooks in public issues. Use
GitHub private vulnerability reporting at
https://github.com/DipanshuHandoo/excelio/security/advisories/new when enabled.
The repository owner should enable private vulnerability reporting before release.
If it is unavailable, contact the owner through GitHub to arrange a private
channel without disclosing exploit details publicly.

Include affected versions, reproduction with synthetic data, impact, and whether
the problem requires specific worker settings or malformed XLSX content. Only the
latest published minor line receives routine security fixes; no response SLA is
promised.

## Untrusted Workbooks

Reading loads a complete XLSX workbook and returns collected rows. Compressed
inputs can expand substantially in memory. Enforce file-size limits, permissions,
request timeouts, and process/container resource limits outside this library.
Workers share a process and are not a security sandbox or memory limit. Cancellation
checks do not interrupt initial parsing or guarantee timely termination.

## Dependency Review

ExcelJS is the sole direct runtime dependency. `npm run audit:dependencies` rejects
high/critical advisories; moderate advisories still require maintainer review and
an explicit release decision. CI does not hold npm publishing credentials.

The setup audit of ExcelJS 4.4.0 reports GHSA-w5hq-g745-h8pq through uuid 8.3.2,
counted as two moderate findings (uuid and its dependent ExcelJS). The advisory
concerns uuid v3/v5/v6 buffer bounds handling. The inspected ExcelJS library uses
uuid v4 in conditional-formatting serialization, not those functions; this does
not establish that all downstream uses or all dependency risks are safe.
No forced downgrade or major-version override has been applied. Review the current
advisory and available upstream fixes before release; audit results may change.