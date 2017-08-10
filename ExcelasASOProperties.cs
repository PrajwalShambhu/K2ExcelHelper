///Author:  Prajwal Shambhu
///Company: K2 Middle East
///Date:    13th January 2016
///Version: v.1.0

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using System.Data.SqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Xml;


namespace K2ME.ExcelAsSO
{
    public class ExcelasASOProperties
    {
        /// <summary>
        /// Describes the functionality of the service object
        /// </summary>
        
         #region private data members

        private int _NoOfColumns;
        private string _UNCPath;
        private Boolean _IsDelete;
        private string _ConnectionString;

        #endregion

        #region ctor(s)

        internal ExcelasASOProperties() { }

        internal ExcelasASOProperties(int NoOfColumns, string UNCPath, Boolean IsDelete, string ConnectionString)
        {
            _NoOfColumns = NoOfColumns;
            _UNCPath = UNCPath;
            _IsDelete = IsDelete;
            _ConnectionString = ConnectionString;
        }

        #endregion


        internal ServiceObject DescribeServiceObject()
        {
            ServiceObject so = new ServiceObject("ExcelDataSmartObject");
            so.Type = "ExcelDataSmartObject";
            so.Active = true;
            so.MetaData.DisplayName = "Excel as SmartObject";
            so.MetaData.Description = "SmartObject returns Excel data";
            so.Properties = DescribeProperties();
            so.Methods = DescribeMethods();
            return so;
        }

        private SourceCode.SmartObjects.Services.ServiceSDK.Objects.Properties DescribeProperties()
        {

            Properties properties = new Properties();

            properties.Add(new Property("ExcelDocumentPath","System.String",SoType.File, new MetaData("ExcelDocumentPath","ExcelDocumentPath")));
            properties.Add(new Property("RowNumber", "System.Int32", SoType.Number, new MetaData("RowNumber", "RowNumber")));
            properties.Add(new Property("ExcelSheetName", "System.String",SoType.Text, new MetaData("ExcelSheetName", "ExcelSheetName")));
            properties.Add(new Property("TableName", "System.String", SoType.Text, new MetaData("TableName", "SmartObjectName")));
            properties.Add(new Property("FilterData", "System.Guid", SoType.Guid, new MetaData("FilterData", "FilterData")));

            for (int i = 1; i <= _NoOfColumns; i++)
            {
                properties.Add(new Property("Column"+i, "System.String", SoType.Text, new MetaData("Column"+i, "Column"+i)));
            }
            return properties;
        }

        private SourceCode.SmartObjects.Services.ServiceSDK.Objects.Methods DescribeMethods()
        {
            SourceCode.SmartObjects.Services.ServiceSDK.Objects.Methods methods = new Methods();
            methods.Add(new Method("ExcelDataList", MethodType.List, new MetaData("ExcelDataList", "Returns a collection Excel sheet."), GetRequiredProperties("ExcelList"), GetMethodParameters(), GetInputProperties("ExcelList"), GetReturnProperties("ExcelList")));
            ////Method to Save File to 
            methods.Add(new Method("SaveExcelData", MethodType.List, new MetaData("SaveExcelData", "Returns a collection Excel sheet."), GetRequiredProperties("SaveData"), GetMethodParameters(), GetInputProperties("SaveData"), GetReturnProperties("SaveData")));
            
            return methods;
        }

        private InputProperties GetInputProperties(string Method)
        {
            InputProperties properties = new InputProperties();
            switch (Method)
            {
                case "ExcelList":
                    #region properties
                    properties.Add(new Property("ExcelDocumentPath", "System.String", SoType.File, new MetaData("ExcelDocumentPath", "Excel Document Path Value")));
                    properties.Add(new Property("RowNumber", "System.Int32", SoType.Number, new MetaData("RowNumber", "RowNumber")));
                    properties.Add(new Property("ExcelSheetName", "System.String", SoType.Text, new MetaData("ExcelSheetName", "Excel Sheet Name")));

                    for (int i = 1; i <= _NoOfColumns; i++)
                    {
                        properties.Add(new Property("Column" + i, "System.String", SoType.Text, new MetaData("Column" + i, "Column" + i)));
                    }
                    break;
                case "SaveData":
                    properties.Add(new Property("ExcelDocumentPath", "System.String", SoType.File, new MetaData("ExcelDocumentPath", "Excel Document Path Value")));
                    properties.Add(new Property("RowNumber", "System.Int32", SoType.Number, new MetaData("RowNumber", "RowNumber")));
                    properties.Add(new Property("ExcelSheetName", "System.String", SoType.Text, new MetaData("ExcelSheetName", "Excel Sheet Name")));
                    properties.Add(new Property("TableName", "System.String", SoType.Text, new MetaData("TableName", "SmartObjectName")));
                    //properties.Add(new Property("FilterData", "System.Guid", SoType.Guid, new MetaData("FilterData", "FilterData")));
                    break;

                    #endregion
            }
            return properties;
        }

        private Validation GetRequiredProperties(string method)
        {
            RequiredProperties properties = new RequiredProperties();
            Validation validation = null;
            validation = new Validation();
            switch (method)
            {
                case "ExcelList":
                    #region properties
                    properties.Add(new Property("ExcelDocumentPath", "System.String", SoType.File, new MetaData("ExcelDocumentPath", "Excel Document Path Value")));
                    properties.Add(new Property("ExcelSheetName", "System.String", SoType.Text, new MetaData("ExcelSheetName", "Excel Sheet Name")));
                    properties.Add(new Property("RowNumber", "System.Int32", SoType.Number, new MetaData("RowNumber", "RowNumber")));
                    break;
                case "SaveData":
                    properties.Add(new Property("ExcelDocumentPath", "System.String", SoType.File, new MetaData("ExcelDocumentPath", "Excel Document Path Value")));
                    properties.Add(new Property("RowNumber", "System.Int32", SoType.Number, new MetaData("RowNumber", "RowNumber")));
                    properties.Add(new Property("ExcelSheetName", "System.String", SoType.Text, new MetaData("ExcelSheetName", "Excel Sheet Name")));
                    properties.Add(new Property("TableName", "System.String", SoType.Text, new MetaData("TableName", "SmartObjectName")));
                    break;

                    #endregion
            }
            validation.RequiredProperties = properties;
            return validation;
        }

        private MethodParameters GetMethodParameters()
        {
            MethodParameters parameters = new MethodParameters();
            return parameters;
        }

        private ReturnProperties GetReturnProperties(string method)
        {
            ReturnProperties properties = new ReturnProperties();
            switch (method)
            {
                case "ExcelList":
                    for (int i = 1; i <= _NoOfColumns; i++)
                    {
                        properties.Add(new Property("Column" + i, "System.String", SoType.Text, new MetaData("Column" + i, "Column" + i)));
                    }
                    break;
                case "SaveData":
                    properties.Add(new Property("FilterData", "System.Guid", SoType.Guid, new MetaData("FilterData", "FilterData")));
                    break;
            }
            return properties;
        }

        //result table returns all Excel data objects
        private DataTable GetResultTable()
        {
            DataTable result = new DataTable();
            for (int i = 1; i <= _NoOfColumns; i++)
            {
                result.Columns.Add("Column"+i, typeof(string));
            }
            return result;
        }

        public DataTable ReadExcelData(Dictionary<string, object> properties, Dictionary<string, object> parameters)
        {
            string ExcelfilePath = properties["ExcelDocumentPath"].ToString();
            string UNCPathValue = _UNCPath; 
            string SheetName = properties["ExcelSheetName"].ToString();
            int StartingRow = Convert.ToInt32(properties["RowNumber"]);
            try
            {
                DataTable ExceldataTable = new DataTable();
                string content = string.Empty;
                string fileName = string.Empty;
                System.Guid guid = Guid.NewGuid();
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.LoadXml(ExcelfilePath);
                XmlNodeList xmlnList = xmldoc.SelectNodes("/file");
                foreach (XmlNode xn in xmlnList)
                {
                    content = xn["content"].InnerText;
                    fileName = xn["name"].InnerText;
                }
                FileStream ar = new FileStream(UNCPathValue+ fileName, System.IO.FileMode.Create);
                ar.Close();
                File.WriteAllBytes(UNCPathValue +guid+ fileName, Convert.FromBase64String(content));

                using (ExcelPackage pck = new ExcelPackage())
                {
                    Stream filestream = File.OpenRead(UNCPathValue +guid+ fileName);
                    pck.Load(filestream);
                    ExcelWorksheet eWS = pck.Workbook.Worksheets[SheetName];
                    ExceldataTable = WorksheetToDataTable(eWS, StartingRow);
                    FileDelete(UNCPathValue + guid + fileName, _IsDelete);
                }
                return ExceldataTable;
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }

        public DataTable WorksheetToDataTable(ExcelWorksheet ws, int firstRowNumber)
        {
            try
            {
                DataTable dt = GetResultTable();
                int totalCols = ws.Dimension.End.Column;
                int totalRows = ws.Dimension.End.Row;
                int startRow = firstRowNumber;
                ExcelRange wsRow;
                DataRow dr;

                for (int rowNum = startRow; rowNum <= totalRows; rowNum++)
                {
                    wsRow = ws.Cells[rowNum, 1, rowNum, _NoOfColumns];
                    dr = dt.NewRow();
                    foreach (var cell in wsRow)
                    {
                        
                        dr[cell.Start.Column - 1] = cell.Text;
                    }

                    dt.Rows.Add(dr);
                }
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void FileDelete(string filePath, Boolean IsDelete)
        {
            if (IsDelete)
            {
                try
                {
                    System.IO.File.Delete(filePath);
                }
                    
                catch(Exception ex)
                {
                }
            }
            
        }

        public DataTable SaveData2Table( Dictionary<string, object> properties, Dictionary<string, object> parameters)
        {
            DataTable ExceldataTable = ReadExcelData(properties, parameters);
            DataTable guidTable = new DataTable();
            DataRow drguid;
           
            System.Guid guid = Guid.NewGuid();
            
            try
            {
                string sql1 = string.Empty;
                string sql = string.Empty;
               
                SqlConnection SqlConnectionObj = new SqlConnection(_ConnectionString);
                SqlCommand cmd = new SqlCommand();
                SqlConnectionObj.Open();
                for (int i = 0; i < ExceldataTable.Rows.Count; i++)
                {
                    sql = "INSERT INTO " + properties["TableName"].ToString() + " (Col1,Col2,Col3,Col4,Col5,Col6,Col7,Col8,Col9,Col10,Col11,Col12,Col13,Col14,Col15,Col16,Col17,Col18,Col19,Col20,Col21,Col22,Col23,Col24,Col25,Col26,Col27,Col28,Col29,Col30, GUIDKey) VALUES ('";
                    sql1 = string.Empty;
                    for (int j = 1; j <= 30; j++)
                    {
                        sql1 =  sql1 + ExceldataTable.Rows[i]["Column" + j].ToString().Trim() + "','";
                    }
                    sql = sql + sql1  + guid + "')";
                    cmd.CommandText = sql;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = SqlConnectionObj;
                    
                    cmd.ExecuteNonQuery();
                }
                SqlConnectionObj.Dispose();
                drguid = guidTable.NewRow();
                guidTable.Columns.Add("FilterData");
                
                drguid["FilterData"] = guid;
                guidTable.Rows.Add(drguid);
                
                return guidTable;
                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
