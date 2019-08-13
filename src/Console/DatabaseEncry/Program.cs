using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseEncry
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectString = @"Server=ALBERTAY\sql2016; Database=BusinessData;User Id=sa;Password=123456";

            using (SqlConnection conn = new SqlConnection(connectString))
            {
                conn.Open();

                string sql = @" INSERT INTO dbo.SearchData
                                            (
                                                CreationTime,
                                                CreatorUserId,
                                                LastModificationTime,
                                                LastModifierUserId,
                                                IsDeleted,
                                                DeleterUserId,
                                                DeletionTime,
                                                Smiles,
                                                Count,
                                                StartTime,
                                                EndTime,
                                                Status,
                                                OrderIndex,
                                                UserId,
                                                Parameter,
                                                ParameterValue
                                            )
                                            VALUES
                                            (SYSDATETIME(), --CreationTime - datetime2(7)
                                                0, --CreatorUserId - bigint
                                                SYSDATETIME(), --LastModificationTime - datetime2(7)
                                                0, --LastModifierUserId - bigint
                                                0, --IsDeleted - bit
                                                0, --DeleterUserId - bigint
                                                SYSDATETIME(), --DeletionTime - datetime2(7)
                                                'NNC1=C(C=CC=C1)N(=O)=O', --Smiles - varchar(900)
                                                0, --Count - int
                                                SYSDATETIME(), --StartTime - datetime2(7)
                                                SYSDATETIME(), --EndTime - datetime2(7)
                                                0, --Status - int
                                                0, --OrderIndex - int
                                                NEWID(), --UserId - uniqueidentifier
                                                'quickly', --Parameter - varchar(256)
                                                N'' -- ParameterValue - nvarchar(max)
                                                )";
                SqlCommand command = conn.CreateCommand();
                command.CommandText = sql;
                int count = command.ExecuteNonQuery();











            }
        }
    }
}
