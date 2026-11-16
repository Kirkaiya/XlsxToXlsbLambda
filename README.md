# XlsxToXlsbLambda

AWS Lambda function (.NET 10) that converts an Excel `.xlsx` workbook to the Excel binary format
(`.xlsb`), using the MIT-licensed [OfficeIMO.Excel](https://github.com/EvotecIT/OfficeIMO) package.
No Excel, COM or paid library is needed. It was built so the Evolve storage agent can turn the
ecs-etl-toolkit OLA export (`.xlsx`) into the `.xlsb` the Empower dashboard imports.

## Request

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

## Response

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

The tests build a synthetic workbook with every construct listed above, so the repo carries no
customer data. To also convert a real OLA export cell for cell, point `OLA_SAMPLE_XLSX` at one:

```powershell
$env:OLA_SAMPLE_XLSX = "C:\path\to\Evolve-OLA-Export.xlsx"; dotnet test XlsxToXlsbLambda.slnx
```

Otherwise that test is reported as skipped.

## Deploy

`src/aws-lambda-tools-defaults.json` sets `dotnet10`, x86_64, **2048 MB** and a **60 s** timeout in
`us-east-1`, the agent's region. Lambda allocates CPU in proportion to memory. A cold conversion of
the 22-sheet export took about 7 s on one full core, so at the template's 512 MB (about 0.3 vCPU)
it could come close to the old 30 s timeout.

```bash
dotnet tool install -g Amazon.Lambda.Tools
cd src
dotnet lambda deploy-function XlsxToXlsbLambda --function-role <execution-role-arn>
```

The execution role needs `AWSLambdaBasicExecutionRole`, plus `s3:GetObject` / `s3:PutObject` on the
buckets used in S3 mode. A caller needs `lambda:InvokeFunction` on the function.
