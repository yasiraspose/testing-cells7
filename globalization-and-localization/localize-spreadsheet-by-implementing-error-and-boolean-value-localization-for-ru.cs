using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook with Russian culture settings
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CultureInfo = new CultureInfo("ru-RU"); // optional, influences number/date parsing
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Apply custom globalization settings that localize Boolean and error values to Russian
        workbook.Settings.GlobalizationSettings = new RussianGlobalizationSettings();

        // Save the localized workbook
        workbook.Save("output.xlsx");
    }

    // Custom globalization settings for Russian language
    private class RussianGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values
        public override string GetBooleanValueString(bool value)
        {
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize Excel error strings
        public override string GetErrorValueString(string error)
        {
            switch (error)
            {
                case "#NAME?":   return "#ИМЯ?";
                case "#DIV/0!":  return "#ДЕЛ/0!";
                case "#REF!":    return "#ССЫЛКА!";
                case "#VALUE!":  return "#ЗНАЧ!";
                case "#N/A":     return "#Н/Д";
                case "#NUM!":    return "#ЧИСЛО!";
                case "#NULL!":   return "#ПУСТО!";
                default:         return base.GetErrorValueString(error);
            }
        }
    }
}