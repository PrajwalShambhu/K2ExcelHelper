///author:  Prajwal Shambhu
///Company: K2 Middle East
///Date:    13th January 2016
///Version: v.1.0

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using System.Data;

namespace K2ME.ExcelAsSO
{
    class ExcelasSOBroker : ServiceAssemblyBase
    {

        #region private data members

        private int _NoOfColumns;
        private string _UNCPath;
        private Boolean _IsDelete;
        private string _ConnectionString;
        
        #endregion


        public override string GetConfigSection()
        {
            base.Service.ServiceConfiguration.Add("No.ofColumns", true, 20);
            base.Service.ServiceConfiguration.Add("UNCPath", true, "c:\\");
            base.Service.ServiceConfiguration.Add("IsDelete", true, true);
            base.Service.ServiceConfiguration.Add("ConnectionString", true, "Integrated=True;IsPrimaryLogin=True;Authenticate=True;EncryptedPassword=False;Host=localhost;Port=5555");
            return base.GetConfigSection();
        }

        public override string DescribeSchema()
        {
            base.Service.Name = "ExcelasSMO";
            base.Service.MetaData.DisplayName = "Excel as SMO";
            base.Service.MetaData.Description = "";

            ExcelasASOProperties wrkProperties = new ExcelasASOProperties();
            base.Service.ServiceObjects.Add(wrkProperties.DescribeServiceObject());

            return base.DescribeSchema();
        }

        public override void Extend()
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override void Execute()
        {
            ValidateConfigSection();
            base.ServicePackage.ResultTable = null;
            DataTable result = new DataTable("Result");
            //System.Guid filterdata;
            try
            {

                foreach (ServiceObject serviceObj in base.Service.ServiceObjects)
                {
                    foreach (Method method in serviceObj.Methods)
                    {
                        Dictionary<string, object> properties = new Dictionary<string, object>();
                        foreach (Property property in serviceObj.Properties)
                        {
                            if ((property.Value != null) && (!string.IsNullOrEmpty(property.Value.ToString())))
                            {
                                properties.Add(property.Name, property.Value);
                            }
                        }

                        // build the method parameters collection
                        Dictionary<string, object> parameters = new Dictionary<string, object>();
                        foreach (MethodParameter parameter in method.MethodParameters)
                        {
                            if ((parameter.Value != null) && (!string.IsNullOrEmpty(parameter.Value.ToString())))
                            {
                                parameters.Add(parameter.Name, parameter.Value);
                            }
                        }

                        if (serviceObj.Name.ToString() == "ExcelDataSmartObject")
                        {
                            if (method.Name == "ExcelDataList")
                            {
                                result = ExcelListData(properties, parameters);
                            }
                            if (method.Name == "SaveExcelData")
                            {
                                result = SaveData(properties, parameters);
                            }
                        }
                    }
                    base.ServicePackage.ResultTable = result;
                                        
                    base.ServicePackage.IsSuccessful = true;
                }
            }
            catch (Exception ex)
            {
                base.ServicePackage.IsSuccessful = false;
                base.ServicePackage.ServiceMessages.Add(new ServiceMessage(ex.Message, MessageSeverity.Error));
            }

        }

        private void ValidateConfigSection()
        {
            ServiceConfiguration config = base.Service.ServiceConfiguration;
            _NoOfColumns = Convert.ToInt32(config["No.ofColumns"].ToString());
            _UNCPath = config["UNCPath"].ToString();
            _ConnectionString = config["ConnectionString"].ToString();
            _IsDelete = Convert.ToBoolean(config["IsDelete"]);


            if (_NoOfColumns == 0)
            {
                base.ServicePackage.IsSuccessful = false;
                base.ServicePackage.ServiceMessages.Add(new ServiceMessage("Please provide number of columns", MessageSeverity.Error));
            }
            if (string.IsNullOrEmpty(_UNCPath))
            {
                base.ServicePackage.IsSuccessful = false;
                base.ServicePackage.ServiceMessages.Add(new ServiceMessage("Please provide UNC Path details", MessageSeverity.Error));
            }
            if (string.IsNullOrEmpty(_ConnectionString))
            {
                base.ServicePackage.IsSuccessful = false;
                base.ServicePackage.ServiceMessages.Add(new ServiceMessage("Please provide K2 Server connection  string details.",MessageSeverity.Error));
            }
            
        }
        
        private DataTable ExcelListData(Dictionary<string, object> properties, Dictionary<string, object> parameters)
        {
            DataTable result;
            ExcelasASOProperties _excelProperties = new ExcelasASOProperties(_NoOfColumns,_UNCPath, _IsDelete, _ConnectionString);
            result = _excelProperties.ReadExcelData(properties, parameters);
            return result;
        }

        private DataTable SaveData(Dictionary<string, object> properties, Dictionary<string, object> parameters)
        {
            DataTable result;
            ExcelasASOProperties _excelProperties = new ExcelasASOProperties(_NoOfColumns, _UNCPath, _IsDelete, _ConnectionString);
            result = _excelProperties.SaveData2Table(properties, parameters);
            return result;
        }
    }
}

