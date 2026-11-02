using Microsoft.Data.SqlClient;

namespace A246FOmmApi.Data
{
    public sealed class SqlConnectionFactory
    {
        private readonly string _connectionString;

        public SqlConnectionFactory(IConfiguration configuration)
        {
            _connectionString =
                configuration.GetConnectionString("A246FDatabase")
                ?? throw new InvalidOperationException(
                    "Connection string 'A246FDatabase' is missing.");
        }

        public SqlConnection CreateConnection()
        {
            return new SqlConnection(_connectionString);
        }
    }
}