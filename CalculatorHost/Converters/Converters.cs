using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace CalculatorHost.Converters;

public class BoolToVisibilityConverter : IValueConverter {
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) {
        return value is true ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) {
        return value is Visibility.Visible;
    }
}

public class InverseBoolToVisibilityConverter : IValueConverter {
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) {
        return value is true ? Visibility.Collapsed : Visibility.Visible;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) {
        return value is not Visibility.Visible;
    }
}

public class NullToVisibilityConverter : IValueConverter {
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) {
        if (value is null) return Visibility.Collapsed;
        if (value is string text && string.IsNullOrWhiteSpace(text)) return Visibility.Collapsed;

        return Visibility.Visible;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) {
        throw new NotSupportedException();
    }
}

public class BoolToStringConverter : IValueConverter {
    private string TrueValue { get; } = "True";
    private string FalseValue { get; } = "False";

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) {
        return value is true ? TrueValue : FalseValue;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) {
        throw new NotSupportedException();
    }
}