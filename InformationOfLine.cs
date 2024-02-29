using NPOI.XWPF.UserModel;

namespace PdfToXls
{
    public class InformationOfLine
    {
        public bool HasValues => !string.IsNullOrWhiteSpace(Date) && !string.IsNullOrWhiteSpace(Description) && !string.IsNullOrWhiteSpace(Value) && !string.IsNullOrWhiteSpace(TotalValue);
        public string Date { get; set; }
        public string Description { get; set; }
        public string Value { get; set; }
        public string TotalValue { get; set; }

        public InformationOfLine()
        {
            Date = string.Empty;
            Description = string.Empty;
            Value = string.Empty;
            TotalValue = string.Empty;
        }

        public void Clear()
        {
            Date = string.Empty;
            Description = string.Empty;
            Value = string.Empty;
            TotalValue = string.Empty;
        }

    }
}