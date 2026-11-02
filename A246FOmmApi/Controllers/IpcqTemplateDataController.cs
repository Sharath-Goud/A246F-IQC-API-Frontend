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

                // Validate Excel measurement header
                if (!string.Equals(
                        ReadCell(7, 2),
                        "Content",
                        StringComparison.OrdinalIgnoreCase))
                {
                    return BadRequest(new
                    {
                        success = false,
                        fileName = uploadedFileName,
                        message = "Unexpected Excel layout. Content header not found in B7."
                    });
                }

                var measurements = new List<string>();

                string parameter = string.Empty;

                decimal nominalValue = 0;
                decimal upperTolerance = 0;
                decimal lowerTolerance = 0;

                int firstMeasurementRow = 0;

                int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;

                if (lastRow < 8)
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "No measurement records found."
                    });
                }

                // Read measurement values and first measurement specifications
                for (int row = 8; row <= lastRow; row++)
                {
                    string content = ReadCell(row, 2);
                    string actual = ReadCell(row, 3);

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

                    // Read specifications from the first valid measurement row
                    if (firstMeasurementRow == 0)
                    {
                        parameter = content;
                        firstMeasurementRow = row;

                        if (!decimal.TryParse(
                            ReadCell(row, 4),
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out nominalValue))
                        {
                            return BadRequest(new
                            {
                                success = false,
                                message = $"Invalid Nominal value at row {row}."
                            });
                        }

                        if (!decimal.TryParse(
                            ReadCell(row, 6),
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out upperTolerance))
                        {
                            return BadRequest(new
                            {
                                success = false,
                                message = $"Invalid UpTol value at row {row}."
                            });
                        }

                        if (!decimal.TryParse(
                            ReadCell(row, 7),
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out lowerTolerance))
                        {
                            return BadRequest(new
                            {
                                success = false,
                                message = $"Invalid LowTol value at row {row}."
                            });
                        }
                    }

                    measurements.Add(
                        actualValue.ToString(
                            CultureInfo.InvariantCulture));
                }

                if (measurements.Count != 32)
                {
                    return BadRequest(new
                    {
                        success = false,
                        fileName = uploadedFileName,
                        message =
                            $"Expected 32 measurements but found {measurements.Count}. " +
                            "The file was not saved."
                    });
                }

                decimal uslValue = nominalValue + upperTolerance;
                decimal lslValue = nominalValue + lowerTolerance;

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
                    // Find exactly one matching template.
                    // Do not select an arbitrary template if duplicates exist.
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
                        [Unique Identifier]
                    FROM dbo.IPQC_TEMPLATE_A246F
                    WHERE
                        REPLACE(UPPER(LTRIM(RTRIM(PART))), ' ', '') = @PartKey
                        AND UPPER(LTRIM(RTRIM(STATION_ID))) = @Station
                        AND ABS(TRY_CONVERT(DECIMAL(18,6), NOMINAL) - @Nominal) < 0.000001
                        AND ABS(TRY_CONVERT(DECIMAL(18,6), USL) - @USL) < 0.000001
                        AND ABS(TRY_CONVERT(DECIMAL(18,6), LSL) - @LSL) < 0.000001;
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
                    string? samplingPlan;
                    string? nominal;
                    string? usl;
                    string? lsl;
                    string? uniqueIdentifier;

                    await using (var command =
                        new SqlCommand(templateSql, connection, transaction))
                    {
                        string partKey = partName
                        .Replace(" ", "")
                        .Trim()
                        .ToUpperInvariant();

                        command.Parameters.Add(
                            "@PartKey",
                            SqlDbType.VarChar,
                            100).Value = partKey;

                        command.Parameters.Add(
                            "@Station",
                            SqlDbType.VarChar,
                            100).Value = machine.Trim().ToUpperInvariant();

                        var nominalParam = command.Parameters.Add(
                            "@Nominal", SqlDbType.Decimal);
                        nominalParam.Precision = 18;
                        nominalParam.Scale = 6;
                        nominalParam.Value = nominalValue;

                        var uslParam = command.Parameters.Add(
                            "@USL", SqlDbType.Decimal);
                        uslParam.Precision = 18;
                        uslParam.Scale = 6;
                        uslParam.Value = uslValue;

                        var lslParam = command.Parameters.Add(
                            "@LSL", SqlDbType.Decimal);
                        lslParam.Precision = 18;
                        lslParam.Scale = 6;
                        lslParam.Value = lslValue;

                        await using var reader =
                            await command.ExecuteReaderAsync(cancellationToken);

                        if (!await reader.ReadAsync(cancellationToken))
                        {
                            return BadRequest(new
                            {
                                success = false,
                                message =
                                    $"No template found for Part='{partName}', " +
                                    $"Parameter='{parameter}', Station='{machine}'. " +
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
                        samplingPlan = GetNullableString(reader, 12);
                        nominal = GetNullableString(reader, 13);
                        usl = GetNullableString(reader, 14);
                        lsl = GetNullableString(reader, 15);
                        uniqueIdentifier = GetNullableString(reader, 16);

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

                        AddNVarChar(null, 10); 

                        AddNVarChar(project, 100);
                        AddNVarChar(site, 100);
                        AddNVarChar(build, 50);
                        AddNVarChar(inspectionStep, 100);
                        AddNVarChar(station, 100);

                        AddNVarChar(templatePart ?? partName, 150);
                        AddNVarChar(vendor ?? company, 150);

                        AddNVarChar(null, 100); 
                        AddNVarChar(null, 50);  

                        AddNVarChar(process, 100);
                        AddNVarChar(category, 100);
                        AddNVarChar(templateParameter ?? parameter, 150);
                        AddNVarChar(uniqueIdentifier, 250);

                        AddNVarChar(unit, 20);
                        AddNVarChar(samplingPlan, 50);

                        AddNVarChar(nominal, 50);
                        AddNVarChar(usl, 50);
                        AddNVarChar(lsl, 50);

                        AddDateTime(reportTime);

                        foreach (string measurement in measurements)
                        {
                            AddNVarChar(measurement, 50);
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