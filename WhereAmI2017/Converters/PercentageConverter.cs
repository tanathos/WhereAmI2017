using System;
using System.ComponentModel;
using System.Globalization;

namespace WhereAmI2017.Converters
{
    internal class PercentageConverter : DoubleConverter
    {
        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
        {
            double parsed = 1;

            Double.TryParse(value.ToString(), out parsed);
            if (parsed > 1)
                value = 1;

            if (parsed < 0)
                value = 0;

            return base.ConvertTo(context, culture, value, destinationType);
        }
    }
}
