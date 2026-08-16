# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | :white_check_mark: |

## Threat Model & Supply Chain Posture

ReverseXL parses untrusted, complex binary and XML-based spreadsheet formats (`.xlsx`, `.xlsm`, `.xls`). To protect downstream enterprise systems from supply-chain risks:

1. **Minimized Attack Surface:** We deliberately restrict third-party libraries exclusively to **Apache POI 5.2.4**.
2. **Safe XML Configuration:** Apache POI 5.2.4 enforces secure XML entity expansion limits by default, mitigating XML External Entity (XXE) and billion laughs attack vectors.
3. **No Dynamic Execution:** Macros and VBA code are inspected statically as string tokens; no embedded code or macros are executed during schema extraction.

## Reporting a Vulnerability

If you discover a potential vulnerability or security issue, please contact the maintainers directly or open a confidential security advisory.
