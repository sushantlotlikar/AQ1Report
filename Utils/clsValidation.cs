namespace AQ1Report.Utils
{
    class clsValidation
    {
        public static bool IsInt32(string value)
        {
            int output;
            return int.TryParse(value, out output);
        }

        public static bool IsNumeric(string value)
        {
            return Microsoft.VisualBasic.Information.IsNumeric(value);
        }
    }
}
