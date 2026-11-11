using ClosedXML.Excel;
using A246FOmmApi.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;

namespace A246FOmmApi.Controllers
{
    [ApiController]
    [Route("api/open/ipqc-template-data-a246f")]
    public class IpqcTemplateDataController : ControllerBase
    {
        private readonly SqlConnectionFactory _connectionFactory;
        private readonly ILogger<IpqcTemplateDataController> _logger;

        public IpqcTemplateDataController(
            SqlConnectionFactory connectionFactory,
            ILogger<IpqcTemplateDataController> logger)
        {
            _connectionFactory = connectionFactory;
            _logger = logger;
        }

        [HttpPost("process-report")]
        [RequestSizeLimit(20_000_000)]
        public async Task<IActionResult> ProcessReport(
            IFormFile file,
            CancellationToken cancellationToken)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new
                {
                    success = false,
                    message = "Please select an Excel file to upload."
                });
            }

            string uploadedFileName = Path.GetFileName(file.FileName);

            var extension = Path.GetExtension(file.FileName);

            if (!string.Equals(extension, ".xlsx",
                    StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new
                {
                    success = false,
                    message = "Only .xlsx files are accepted."
                });
            }

            try
            {
                using var stream = new MemoryStream();

                await file.CopyToAsync(stream, cancellationToken);
                stream.Position = 0;

                using var workbook = new XLWorkbook(stream);

                var sheet = workbook.Worksheets.FirstOrDefault();

                if (sheet == null)
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "The Excel file contains no worksheet."
                    });
                }

                string ReadCell(int row, int column) =>
                    sheet.Cell(row, column).GetFormattedString().Trim();

                // Read the SPC A header information.
                string reportName = ReadCell(1, 2);
                string partName = ReadCell(2, 2);
                string partNumber = ReadCell(3, 2);
                string operation = ReadCell(4, 2);
                string company = ReadCell(5, 2);
                string machine = ReadCell(5, 5);
                string reportTimeText = ReadCell(1, 7);
                string parameter = Path.GetFileNameWithoutExtension(uploadedFileName).Trim();

                if (string.IsNullOrWhiteSpace(partName))
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "Part Name is missing from cell B2."
                    });
                }

                if (string.IsNullOrWhiteSpace(machine))
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "Machine is missing from cell E5."
                    });
                }

                DateTime reportTime;

                string[] acceptedFormats =
                {
                    "yyyy/MM/dd--HH:mm",
                    "yyyy/MM/dd HH:mm",
                    "yyyy-MM-dd HH:mm",
                    "yyyy-MM-ddTHH:mm:ss",
                    "yyyy/MM/dd--HH:mm:ss"
                };

                bool validTime = DateTime.TryParseExact(
                    reportTimeText,
                    acceptedFormats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out reportTime);

                if (!validTime)
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = $"Invalid report time: {reportTimeText}. " +
                                  "Expected format: yyyy/MM/dd--HH:mm"
                    });
                }


                // ==========================================
                // DYNAMIC EXCEL VALIDATION
                // Based on standard FAI 5-2.xlsx headers
                // ==========================================

                int headerRow = 7;
                int dataStartRow = 8;

                // STEP 1: Identify all Excel columns by header name

                var headerColumns = new Dictionary<string, int>(
                    StringComparer.OrdinalIgnoreCase);

                int lastColumn = sheet.Row(headerRow)
                    .LastCellUsed()?.Address.ColumnNumber ?? 0;

                for (int col = 1; col <= lastColumn; col++)
                {
                    string header = ReadCell(headerRow, col).Trim();

                    if (string.IsNullOrWhiteSpace(header))
                        continue;

                    if (headerColumns.ContainsKey(header))
                    {
                        return BadRequest(new
                        {
                            success = false,
                            message = $"Duplicate header found: {header}"
                        });
                    }

                    headerColumns.Add(header, col);
                }

                // STEP 2: Required standard columns

                string[] requiredHeaders =
                {
    "Content",
    "Actual",
    "Nominal",
    "UpTol",
    "LowTol"
};

                var missingHeaders = requiredHeaders
                    .Where(h => !headerColumns.ContainsKey(h))
                    .ToList();

                if (missingHeaders.Count > 0)
                {
                    return BadRequest(new
                    {
                        success = false,
                        fileName = uploadedFileName,
                        message = "Missing required columns: " +
                                  string.Join(", ", missingHeaders)
                    });
                }

                // STEP 3: Get dynamic column positions

                int contentCol = headerColumns["Content"];
                int actualCol = headerColumns["Actual"];
                int nominalCol = headerColumns["Nominal"];
                int upTolCol = headerColumns["UpTol"];
                int lowTolCol = headerColumns["LowTol"];

                // STEP 4: Initialize measurement variables

                var measurements = new List<string>();

                string? samplingPlan = null;
                decimal nominalValue = 0;
                decimal upperTolerance = 0;
                decimal lowerTolerance = 0;

                int firstMeasurementRow = 0;

                int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;

                if (lastRow < dataStartRow)
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "No measurement records found."
                    });
                }

                // STEP 5: Read measurement data

                for (int row = dataStartRow; row <= lastRow; row++)
                {
                    string content = ReadCell(row, contentCol);
                    string actual = ReadCell(row, actualCol);

                    if (string.IsNullOrWhiteSpace(content))
                        continue;

                    if (content.StartsWith(
                        "SPC A #",
                        StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (!decimal.TryParse(
                        actual,
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out decimal actualValue))
                        continue;

                    // Read specification values from first valid row

                    if (firstMeasurementRow == 0)
                    {
                        firstMeasurementRow = row;

                        string nominalText = ReadCell(row, nominalCol);
                        string upTolText = ReadCell(row, upTolCol);
                        string lowTolText = ReadCell(row, lowTolCol);

                        if (!decimal.TryParse(
                            nominalText,
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out nominalValue))
                        {
                            return BadRequest(new
                            {
                                success = false,
                                message = $"Invalid Nominal value at row {row}: {nominalText}"
                            });
                        }

                        if (!decimal.TryParse(
                            upTolText,
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out upperTolerance))
                        {
                            return BadRequest(new
                            {
                                success = false,
                                message = $"Invalid UpTol value at row {row}: {upTolText}"
                            });
                        }

                        if (!decimal.TryParse(
                            lowTolText,
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out lowerTolerance))
                        {
                            return BadRequest(new
                            {
                                success = false,
                                message = $"Invalid LowTol value at row {row}: {lowTolText}"
                            });
                        }
                    }

                    measurements.Add(
                        actualValue.ToString(
                            CultureInfo.InvariantCulture));
                }


                if (measurements.Count < 1 || measurements.Count > 32)
                {
                    return BadRequest(new
                    {
                        success = false,
                        fileName = uploadedFileName,
                        message =
                            $"Expected between 1 and 32 measurements, but found {measurements.Count}. " +
                            "The file was not saved."
                    });
                }

                samplingPlan = $"{measurements.Count}x/Lot";
                decimal uslValue = upperTolerance;
                decimal lslValue = lowerTolerance;

                if (string.IsNullOrWhiteSpace(parameter))
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "No valid measurement parameter was found."
                    });
                }

                await using var connection =
                    _connectionFactory.CreateConnection();

                await connection.OpenAsync(cancellationToken);

                await using var transaction =
                    (SqlTransaction)await connection.BeginTransactionAsync(
                        cancellationToken);

                try
                {
                    const string templateSql = """
                    SELECT
                        TEMPLATE_ID,
                        PROJECT,
                        SITE,
                        BUILD,
                        INSPECTION_STEP,
                        STATION_ID,
                        PART,
                        VENDOR,
                        PROCESS,
                        CATEGORY,
                        PARAMETER,
                        UNIT,
                        SAMPLING_PLAN,
                        NOMINAL,
                        USL,
                        LSL,
                        [Unique Identifier],
                        WOM,
                        LOT_ID,
                        LINE_ID
                    FROM [master].[dbo].[IPQC_TEMPLATE_A246F]
                    WHERE
                        UPPER(LTRIM(RTRIM(PART))) = @Part
                        AND UPPER(LTRIM(RTRIM(VENDOR))) = @Vendor
                        AND REPLACE(
                            UPPER(LTRIM(RTRIM(PARAMETER))),
                            '/',
                            ''
                        ) = @Parameter
                    """;
                    long templateId;
                    string? project;
                    string? site;
                    string? build;
                    string? inspectionStep;
                    string station;
                    string? templatePart;
                    string? vendor;
                    string? process;
                    string? category;
                    string? templateParameter;
                    string? unit;
                    string? nominal;
                    string? usl;
                    string? lsl;
                    string? uniqueIdentifier;
                    string? wom;
                    string? lotId;
                    string? lineId;

                    await using (var command =
                        new SqlCommand(templateSql, connection, transaction))
                    {
                        string partValue = partName.Trim().ToUpperInvariant();

                        string vendorValue = company.Trim().ToUpperInvariant();

                        string parameterValue = parameter
                            .Trim()
                            .ToUpperInvariant()
                            .Replace("/", "");

                        command.Parameters.Add("@Part", SqlDbType.VarChar, 150).Value = partValue;
                        command.Parameters.Add("@Vendor", SqlDbType.VarChar, 150).Value = vendorValue;
                        command.Parameters.Add("@Parameter", SqlDbType.VarChar, 150).Value = parameterValue;

                        await using var reader =
                            await command.ExecuteReaderAsync(cancellationToken);

                        if (!await reader.ReadAsync(cancellationToken))
                        {
                            return BadRequest(new
                            {
                                success = false,
                                message =
                                    $"No template found for Part='{partName}', " +
                                    $"Vendor='{company}', Parameter='{parameter}'. " +
                                    "No report was saved."
                            });
                        }

                        templateId = reader.GetInt64(0);
                        project = GetNullableString(reader, 1);
                        site = GetNullableString(reader, 2);
                        build = GetNullableString(reader, 3);
                        inspectionStep = GetNullableString(reader, 4);
                        station = reader.GetString(5);
                        templatePart = GetNullableString(reader, 6);
                        vendor = GetNullableString(reader, 7);
                        process = GetNullableString(reader, 8);
                        category = GetNullableString(reader, 9);
                        templateParameter = GetNullableString(reader, 10);
                        unit = GetNullableString(reader, 11);
                        //samplingPlan = GetNullableString(reader, 12);
                        //nominal = GetNullableString(reader, 13);
                        //usl = GetNullableString(reader, 14);
                        //lsl = GetNullableString(reader, 15);
                        uniqueIdentifier = GetNullableString(reader, 16);
                        wom = GetNullableString(reader, 17);
                        lotId = GetNullableString(reader, 18);
                        lineId = GetNullableString(reader, 19);

                        if (await reader.ReadAsync(cancellationToken))
                        {
                            return BadRequest(new
                            {
                                success = false,
                                message =
                                    "Multiple matching templates were found. " +
                                    "Please make the template match unique before importing."
                            });
                        }
                    }

                    const string duplicateSql = """
                        SELECT COUNT_BIG(*)
                        FROM dbo.IPQC_DIMENSION_REPORT_A246F
                        WHERE SourceFileName = @FileName
                          AND SourceRowNumber = @SourceRowNumber;
                        """;

                    await using (var duplicateCommand =
                        new SqlCommand(duplicateSql, connection, transaction))
                    {
                        duplicateCommand.Parameters.Add(
                            "@FileName",
                            SqlDbType.NVarChar,
                            260).Value = uploadedFileName;

                        duplicateCommand.Parameters.Add(
                            "@SourceRowNumber",
                            SqlDbType.Int).Value = firstMeasurementRow;

                        long count = (long)await duplicateCommand.ExecuteScalarAsync(
                            cancellationToken);

                        if (count > 0)
                        {
                            await transaction.RollbackAsync(cancellationToken);

                            return Conflict(new
                            {
                                success = false,
                                message =
                                    "This file has already been imported. " +
                                    "No duplicate report was inserted."
                            });
                        }
                    }

                    var columns = new List<string>
                    {
                        "WOM", "Project", "Site", "Build",
                        "Inspection_Step", "Station_ID", "Part",
                        "Vendor", "Lot_ID", "Line_ID", "Process",
                        "Category", "Parameter", "[Unique Identifier]",
                        "Unit", "Sampling_Plan", "Nominal", "USL",
                        "LSL", "[Time]"
                    };

                    for (int i = 1; i <= 32; i++)
                    {
                        columns.Add($"[{i}]");
                    }

                    columns.Add("SourceFileName");
                    columns.Add("SourceRowNumber");

                    string columnSql = string.Join(", ", columns);
                    string valueSql = string.Join(", ",
                        columns.Select((_, i) => $"@p{i}"));

                    string insertSql = $"""
                        INSERT INTO dbo.IPQC_DIMENSION_REPORT_A246F
                            ({columnSql})
                        VALUES
                            ({valueSql});

                        SELECT CAST(SCOPE_IDENTITY() AS bigint);
                        """;


                    await using (var insertCommand =
                        new SqlCommand(insertSql, connection, transaction))
                    {
                        int p = 0;

                        // Add NVARCHAR parameter
                        void AddNVarChar(object? value, int size)
                        {
                            var sqlParameter = insertCommand.Parameters.Add(
                                $"@p{p++}",
                                SqlDbType.NVarChar,
                                size);

                            sqlParameter.Value = value ?? DBNull.Value;
                        }

                        // Add DATETIME2 parameter
                        void AddDateTime(DateTime value)
                        {
                            insertCommand.Parameters.Add(
                                $"@p{p++}",
                                SqlDbType.DateTime2).Value = value;
                        }

                        // Add INT parameter
                        void AddInt(int value)
                        {
                            insertCommand.Parameters.Add(
                                $"@p{p++}",
                                SqlDbType.Int).Value = value;
                        }

                        AddNVarChar(wom, 50);

                        AddNVarChar(project, 100);
                        AddNVarChar(site, 100);
                        AddNVarChar(build, 50);
                        AddNVarChar(inspectionStep, 100);
                        AddNVarChar(station, 100);

                        AddNVarChar(templatePart ?? partName, 150);
                        AddNVarChar(vendor ?? company, 150);

                        AddNVarChar(partNumber, 100);
                        AddNVarChar(lineId, 50);

                        AddNVarChar(process, 100);
                        AddNVarChar(category, 100);
                        AddNVarChar(templateParameter ?? parameter, 150);
                        AddNVarChar(uniqueIdentifier, 250);

                        AddNVarChar(unit, 20);
                        AddNVarChar(samplingPlan, 50);

                        AddNVarChar(
                            nominalValue.ToString("0.###", CultureInfo.InvariantCulture),
                            50);

                        AddNVarChar(
                            uslValue.ToString("0.###", CultureInfo.InvariantCulture),
                            50);

                        AddNVarChar(
                            lslValue.ToString("0.###", CultureInfo.InvariantCulture),
                            50);

                        AddDateTime(reportTime);


                        for (int i = 0; i < 32; i++)
                        {
                            if (i < measurements.Count)
                            {
                                AddNVarChar(measurements[i], 50);
                            }
                            else
                            {
                                AddNVarChar(null, 50);
                            }
                        }


                        AddNVarChar(uploadedFileName, 260);
                        AddInt(firstMeasurementRow);

                        if (insertCommand.Parameters.Count != 54)
                        {
                            throw new InvalidOperationException(
                                $"Expected 54 SQL parameters but found " +
                                $"{insertCommand.Parameters.Count}.");
                        }

                        long reportId = Convert.ToInt64(
                            await insertCommand.ExecuteScalarAsync(cancellationToken));

                        await transaction.CommitAsync(cancellationToken);

                        return Ok(new
                        {
                            success = true,
                            message = $"Excel file '{uploadedFileName}' saved successfully.",
                            reportId,
                            templateId,
                            fileName = uploadedFileName,
                            measurementCount = measurements.Count,
                            partName,
                            parameter,
                            station,
                            reportTime
                        });
                    }

                }
                catch
                {
                    await transaction.RollbackAsync(cancellationToken);
                    throw;
                }
            }
            catch (OperationCanceledException)
            {
                return StatusCode(408, new
                {
                    success = false,
                    message = "The Excel import was cancelled."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(
                    ex,
                    "Excel import failed for file: {FileName}",
                    uploadedFileName);

                return StatusCode(500, new
                {
                    success = false,
                    fileName = uploadedFileName,
                    message = $"Failed to import '{uploadedFileName}'.",
                    error = ex.Message,
                    innerError = ex.InnerException?.Message
                });
            }
        }

        private static string? GetNullableString(
            SqlDataReader reader,
            int ordinal)
        {
            return reader.IsDBNull(ordinal)
                ? null
                : Convert.ToString(
                    reader.GetValue(ordinal),
                    CultureInfo.InvariantCulture);
        }
    }
}