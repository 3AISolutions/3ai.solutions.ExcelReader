namespace _3ai.solutions.ExcelReader;

public sealed record ExcelData(string Worksheet, string Column, uint Row, string Formula, string TextValue);
