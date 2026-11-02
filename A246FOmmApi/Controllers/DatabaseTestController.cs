using A246FOmmApi.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace A246FOmmApi.Controllers
{
    [ApiController]
    [Route("api/open/database-test")]
    public class DatabaseTestController : ControllerBase
    {
        private readonly SqlConnectionFactory _connectionFactory;

        public DatabaseTestController(
            SqlConnectionFactory connectionFactory)
        {
            _connectionFactory = connectionFactory;
        }

        [HttpGet]
        public async Task<IActionResult> TestConnection(
            CancellationToken cancellationToken)
        {
            try
            {
                await using var connection =
                    _connectionFactory.CreateConnection();

                await connection.OpenAsync(cancellationToken);

                await using var command = connection.CreateCommand();

                command.CommandText = """
                    SELECT
                        DB_NAME() AS DatabaseName,
                        @@SERVERNAME AS ServerName,
                        CASE
                            WHEN OBJECT_ID(
                                'dbo.IPQC_TEMPLATE_A246F', 'U'
                            ) IS NOT NULL
                            THEN 1 ELSE 0
                        END AS TemplateTableExists,
                        CASE
                            WHEN OBJECT_ID(
                                'dbo.IPQC_TEMPLATE_DATA_A246F', 'U'
                            ) IS NOT NULL
                            THEN 1 ELSE 0
                        END AS MeasurementTableExists;
                    """;

                await using var reader =
                    await command.ExecuteReaderAsync(cancellationToken);

                if (!await reader.ReadAsync(cancellationToken))
                {
                    return StatusCode(500, new
                    {
                        success = false,
                        message = "Database verification returned no results."
                    });
                }

                var databaseName = reader.GetString(0);
                var serverName = reader.IsDBNull(1)
                    ? null
                    : reader.GetString(1);

                var templateExists = reader.GetInt32(2) == 1;
                var measurementExists = reader.GetInt32(3) == 1;

                return Ok(new
                {
                    success = true,
                    database = databaseName,
                    server = serverName,
                    templateTableExists = templateExists,
                    measurementTableExists = measurementExists,
                    message = templateExists && measurementExists
                        ? "Database connection and A246F tables verified."
                        : "Connected, but one or more A246F tables are missing."
                });
            }
            catch (SqlException)
            {
                return StatusCode(500, new
                {
                    success = false,
                    message =
                        "SQL Server connection or query failed. " +
                        "Check the server, database, credentials, " +
                        "network access, and database permissions."
                });
            }
            catch (OperationCanceledException)
            {
                return StatusCode(408, new
                {
                    success = false,
                    message = "The database test was cancelled."
                });
            }
        }
    }
}