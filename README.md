# XlsxToXlsb

Converts an Excel `.xlsx` workbook to the Excel binary format (`.xlsb`) with the MIT-licensed
[OfficeIMO.Excel](https://github.com/EvotecIT/OfficeIMO) package. No Excel, COM or paid library is
needed. It was built so the Evolve storage agent can turn the ecs-etl-toolkit OLA export (`.xlsx`)
into the `.xlsb` the Empower dashboard imports.

| Project | What it is |
|---|---|
| `src/XlsxToXlsb.Core` | The conversion (`XlsbConverter`). Shared by both front ends below. |
| `src/xlsx2xlsb` | Command-line tool, published self-contained for **linux-arm64** and bundled in the agent's container image. **This is what the agent uses.** |
| `src/XlsxToXlsbLambda` | AWS Lambda wrapper (.NET 10, x86_64), for callers outside the agent image. Not deployed today. |
| `test` | xunit tests for all three. |

## Command-line tool: `xlsx2xlsb`

```bash
xlsx2xlsb <input.xlsx> <output.xlsb>
xlsx2xlsb --version
```

It prints **one line of JSON** to stdout whether it succeeds or fails, so a caller parses stdout:

```json
{"ok":true,"outputPath":"/tmp/out.xlsb","sizeBytes":130651,"sheetCount":22,"adjustments":{"conditionalFormatsRemoved":52,...}}
{"ok":false,"error":"The input is not an .xlsx workbook (it is not a ZIP package)."}
```

On failure the error also goes to stderr. Exit codes: `0` converted, `1` the workbook cannot be
converted, `2` usage, `3` a file could not be read or written. The `.xlsb` is written to a temporary
file and moved into place, so a failed run never leaves a partial output.

### Publishing for the agent image (linux-arm64)

Bedrock AgentCore runs linux/arm64 only. The `linux-arm64` publish profile makes a self-contained
build (no .NET runtime needed in the image), ReadyToRun (precompiled, so a cold run skips JIT), and
**untrimmed** (OpenXml uses reflection). It uses invariant globalization because the agent's
`python:3.11-slim` base has no libicu, and nothing in the conversion depends on culture data.

```bash
dotnet publish src/xlsx2xlsb/xlsx2xlsb.csproj -p:PublishProfile=linux-arm64
# -> src/xlsx2xlsb/bin/publish/linux-arm64/  (about 146 MB; run ./xlsx2xlsb)
```

The agent's Dockerfile builds it this way in a .NET SDK stage and copies the folder to
`/opt/xlsx2xlsb/`. Cross-compiling from x64 to arm64 works, so no emulated SDK is needed.

## Lambda

The Lambda handler is `XlsxToXlsbLambda::XlsxToXlsbLambda.Function::FunctionHandler`.

### Request

Supply the workbook **inline** or **from S3**, but not both.

```json
{ "fileName": "Evolve-OLA-Acme-Export.xlsx", "xlsxBase64": "UEsDBBQ..." }
```

```json
{
  "sourceBucket": "my-bucket", "sourceKey": "exports/Evolve-OLA-Acme-Export.xlsx",
  "destinationBucket": "my-bucket", "destinationKey": "exports/Evolve-OLA-Acme-Export.xlsb"
}
```

| Field | Meaning |
|---|---|
| `xlsxBase64` | The `.xlsx`, base64-encoded. Keep the file under about 4 MB: a synchronous Lambda payload is capped at 6 MB and base64 adds a third. |
| `sourceBucket`, `sourceKey` | Read the `.xlsx` from S3 instead. |
| `fileName` | Optional. Names the output (`.xlsx` → `.xlsb`). Defaults to the source key, or `workbook.xlsx`. |
| `destinationBucket` | Optional. When set, the `.xlsb` is written to S3 rather than returned inline. Required for an output over about 4.4 MB. |
| `destinationKey` | Optional. Defaults to the source key (or file name) with an `.xlsb` extension. |

### Response

```json
{
  "fileName": "Evolve-OLA-Acme-Export.xlsb",
  "sizeBytes": 130651,
  "sheetCount": 22,
  "xlsbBase64": "UEsDBBQ...",
  "adjustments": {
    "conditionalFormatsRemoved": 52,
    "differentialFormatsRemoved": 74,
    "indexedColorPaletteRemoved": true,
    "styleExtensionsRemoved": false,
    "worksheetsWithColumnRangesNormalized": 13,
    "autoFilterSortStatesRemoved": 6
  }
}
```

With `destinationBucket` set, `xlsbBase64` is replaced by `destinationBucket` and `destinationKey`.
A failure (bad input, or a workbook feature OfficeIMO cannot encode) is thrown, so the caller sees a
Lambda function error whose `errorMessage` says why. Nothing is dropped silently.

### Deploy

`src/XlsxToXlsbLambda/aws-lambda-tools-defaults.json` sets `dotnet10`, x86_64, **2048 MB** and a **60 s** timeout in
`us-east-1`, the agent's region. Lambda allocates CPU in proportion to memory. A cold conversion of
the 22-sheet export took about 7 s on one full core (without ReadyToRun), so at the template's 512 MB (about 0.3 vCPU)
it could come close to the old 30 s timeout.

```bash
dotnet tool install -g Amazon.Lambda.Tools
cd src/XlsxToXlsbLambda
dotnet lambda deploy-function XlsxToXlsbLambda --function-role <execution-role-arn>
```

The execution role needs `AWSLambdaBasicExecutionRole`, plus `s3:GetObject` / `s3:PutObject` on the
buckets used in S3 mode. A caller needs `lambda:InvokeFunction` on the function.

## What the conversion changes

OfficeIMO.Excel 3.6.1's native XLSB writer refuses some constructs it cannot encode yet, and the OLA
export contains all of them. `XlsbConverter` normalizes these in memory before saving. **No cell value
changes.**

| Construct | What happens | Effect on the workbook |
|---|---|---|
| Conditional formatting and its dxf styles | Removed | Rule-based highlighting is gone. Cell values and static cell formatting stay. |
| Indexed-color palette (`<colors>`) | Removed | Excel uses its default palette. The OLA export's palette is the standard one. |
| Stylesheet extensions | Removed (none in the OLA export) | None visible |
| Overlapping column ranges (a range, then one column inside it) | Rewritten as ordered, disjoint ranges; the later definition of a column wins | Column widths, hidden state and outline levels kept |
| Sort state saved inside an AutoFilter | Removed | Filters stay; rows keep their order |

The response's `adjustments` reports what was changed on each call.

## Verification

Tested on `Evolve-OLA-Shutterfly_Smoke-Export.xlsx`, the 22-sheet toolkit OLA export for the
Shutterfly regression input:

- **Values:** all 21,271 non-blank cells are identical in the `.xlsb`, read with an independent reader
  (python-calamine). All 13,208 numeric cells are bit-for-bit equal. The same holds on Linux (Ubuntu,
  .NET 10 runtime) and on Windows.
- **Excel:** opens the output as a native `.xlsb` (`FileFormat` 50). It has 22 sheets in the original
  order, the same 14 defined names (the AutoFilter ranges), and filters still on.
- **Empower:** not yet tested with this output. Upload it once to confirm.

## Build and test

```bash
dotnet test XlsxToXlsbLambda.slnx
```

The tests cover the converter, the CLI (run in-process) and the Lambda handler. They build a synthetic workbook with every construct listed above, so the repo carries no
customer data. To also convert a real OLA export cell for cell, point `OLA_SAMPLE_XLSX` at one:

```powershell
$env:OLA_SAMPLE_XLSX = "C:\path\to\Evolve-OLA-Export.xlsx"; dotnet test XlsxToXlsbLambda.slnx
```

Otherwise that test is reported as skipped.
