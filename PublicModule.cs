using IgalDAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;

namespace Export2ExlEngine
{
    public enum ExportType
    {
        NoType,
        Excel,
        CSV
    }

    internal static class Tools
    {
        public static bool IsWriteable(this DirectoryInfo me)
        {
            AuthorizationRuleCollection rules;
            WindowsIdentity identity;
            try
            {
                rules = me.GetAccessControl().GetAccessRules(true, true, typeof(System.Security.Principal.SecurityIdentifier));
                identity = WindowsIdentity.GetCurrent();
            }
            catch (UnauthorizedAccessException uae)
            {
                Debug.WriteLine(uae.ToString());
                return false;
            }
            bool isAllow = false;
            string userSID = identity.User.Value;
            foreach (FileSystemAccessRule rule in rules)
            {
                if (rule.IdentityReference.ToString() == userSID || identity.Groups.Contains(rule.IdentityReference))
                {
                    if ((rule.FileSystemRights.HasFlag(FileSystemRights.Write) ||
                        rule.FileSystemRights.HasFlag(FileSystemRights.WriteAttributes) ||
                        rule.FileSystemRights.HasFlag(FileSystemRights.WriteData) ||
                        rule.FileSystemRights.HasFlag(FileSystemRights.CreateDirectories) ||
                        rule.FileSystemRights.HasFlag(FileSystemRights.CreateFiles)) && rule.AccessControlType == AccessControlType.Deny)
                        return false;
                    else if ((rule.FileSystemRights.HasFlag(FileSystemRights.Write) &&
                        rule.FileSystemRights.HasFlag(FileSystemRights.WriteAttributes) &&
                        rule.FileSystemRights.HasFlag(FileSystemRights.WriteData) &&
                        rule.FileSystemRights.HasFlag(FileSystemRights.CreateDirectories) &&
                        rule.FileSystemRights.HasFlag(FileSystemRights.CreateFiles)) && rule.AccessControlType == AccessControlType.Allow)
                        isAllow = true;
                }
            }
            return isAllow;
        }

        public static string TempTableFromDataTableCmdText(DataTable dt, string sTableName, SqlFields formatFields = null)
        {
            clsParamField field = new clsParamField();            
             
            StringBuilder sbTempTable = new StringBuilder($"CREATE TABLE {sTableName}(");
            string sColDef = "";
            foreach (DataColumn col in dt.Columns)
            {
                switch (col.DataType.ToString())
                {
                    case "System.Int64":
                        sColDef = $"[{col.ColumnName}] bigint ";
                        sColDef += (col.AutoIncrement) ? " Identity (" + col.AutoIncrementSeed.ToString() + "," + col.AutoIncrementStep.ToString() + ")," : ",";
                        sbTempTable.AppendLine(sColDef);
                        break;
                    case "System.Int32":
                        sColDef = $"[{col.ColumnName}] int ";
                        sColDef += (col.AutoIncrement) ? " Identity (" + col.AutoIncrementSeed.ToString() + "," + col.AutoIncrementStep.ToString() + ")," : ",";
                        sbTempTable.AppendLine(sColDef);
                        break;
                    case "System.DateTime":
                        //igal 13/7/20 - as for any date the system returns DateTime type - check the actual type of data
                        if (formatFields.Count>0 && (field = formatFields.GetItemByFieldName(col.ColumnName)) != null)
                            sbTempTable.AppendLine($"[{col.ColumnName}] {field.sqlFieldType.ToString()},");
                        else
                            sbTempTable.AppendLine($"[{col.ColumnName}] datetime,");                            
                        break;
                    case "System.String":
                        if (col.MaxLength > 8000)                        
                            sColDef = $"[{col.ColumnName}] text, ";                                                    
                        else
                        {
                            sColDef = $"[{col.ColumnName}] varchar( ";
                            switch (col.MaxLength)
                            {
                                case -1:
                                    sColDef += "8000"; //defining varchar(max) won't show data in Microsoft query in excel
                                    break;
                                case 0:
                                    sColDef += "1";
                                    break;
                                default:
                                    sColDef += col.MaxLength.ToString();
                                    break;
                            }                             
                            sColDef += "), ";
                        }
                        sbTempTable.AppendLine(sColDef);
                        break;
                    case "System.Single":
                    case "System.Double":
                        sbTempTable.AppendLine($"[{col.ColumnName}] float , ");                        
                        break;                                                                    
                    case "System.Int16":
                        sbTempTable.AppendLine($"[{col.ColumnName}] smallint , ");
                        break;
                    case "System.Boolean":
                        sbTempTable.AppendLine($"[{col.ColumnName}] bit , ");
                        break;
                    case "System.Decimal":
                        sbTempTable.AppendLine($"[{col.ColumnName}] decimal(19,4) , ");
                        break;
                    case "System.Byte":
                        sbTempTable.AppendLine($"[{col.ColumnName}] tinyint, ");
                        break;
                    case "System.Guid":
                        sbTempTable.AppendLine($"[{col.ColumnName}] uniqueidentifier, ");
                        break;                        
                    default:
                        throw new Exception($"No definition for column type {col.DataType.ToString()}");
                }
            }

            string pks = "";
            if (dt.PrimaryKey.Length > 0)
            {
                pks = "CONSTRAINT PK_" + sTableName + " PRIMARY KEY (";
                for (int i = 0; i < dt.PrimaryKey.Length; i++)
                {
                    pks += dt.PrimaryKey[i].ColumnName + ",";
                }
                pks = pks.Substring(0, pks.Length - 1) + ")";

            }
            if (pks != "")
                sbTempTable.AppendLine(pks);
            else
                sbTempTable.Remove(sbTempTable.Length - 1, 1);
            sbTempTable.Append(")");

            return sbTempTable.ToString();
        }

        public static string GetFieldsSelect(DataTable dt, SqlFields formatFields = null)
        {
            List<string> FieldList = new List<string>();
            clsParamField field = new clsParamField();

            foreach (DataColumn col in dt.Columns)
            {
                switch (col.DataType.ToString())
                {
                    case "System.Guid":
                        FieldList.Add($"CAST([{col.ColumnName}] as varchar(50)) as [{col.ColumnName}]"); //excel doesn't see guid column so make it custom cast for select from temp table
                        break;
                    case "System.DateTime":
                        //igal 23/2/21 - write datetime /wo unnesessary zeroes.
                        //check the actual type of data
                        if (formatFields.Count > 0 && (field = formatFields.GetItemByFieldName(col.ColumnName)) != null && field.sqlFieldType.sqlFieldType == "date")
                            FieldList.Add($"CONVERT(varchar(10), CAST([{col.ColumnName}] as date), 103) as [{col.ColumnName}]"); //date dd/mm/yyyy
                        else
                            FieldList.Add($"CONVERT(varchar(10),[{col.ColumnName}], 103) + ' ' + CONVERT(varchar(8),[{col.ColumnName}], 108) as [{col.ColumnName}]"); //datetime dd/mm/yyyy hh:mm:ss                        
                        break;
                    default:
                        FieldList.Add($"[{col.ColumnName}]");
                        break;
                }
            }
            
            return string.Join(",", FieldList);
        }
    }
}