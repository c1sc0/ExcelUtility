using System;

namespace ExcelUtil
{
    public class ControlNotFoundException : Exception
    {
        public ControlNotFoundException(string message) : base($"Not existing control name: {message}")
        {
        }
    }
}