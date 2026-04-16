using System.IO;
using System.Text.Json;
using CalculatorHost.Models;

namespace CalculatorHost.Services;

/// <summary>
///     Loads macro button configuration from a JSON file placed next to the .xlsm file.
///     File name: macros.json
///     Format:
///     {
///     "Buttons": [
///     { "Label": "Pobierz siły",   "MacroName": "PobierzSily",   "Tooltip": "" },
///     { "Label": "Wstępna geometria", "MacroName": "WstepnaGeometria", "Tooltip": "" }
///     ]
///     }
///     If macros.json does not exist, an empty list is returned — no buttons shown.
/// </summary>
public class MacroConfigService {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNameCaseInsensitive = true
    };

    public static List<MacroButtonConfig> LoadForCalculator(string calculatorFilePath) {
        var directory = Path.GetDirectoryName(calculatorFilePath);
        if (string.IsNullOrEmpty(directory)) return [];

        var jsonPath = Path.Combine(directory, "macros.json");
        if (!File.Exists(jsonPath)) return [];

        try {
            var json = File.ReadAllText(jsonPath);
            var file = JsonSerializer.Deserialize<MacroButtonsFile>(json, JsonOptions);
            return file?.Buttons ?? [];
        }
        catch {
            return [];
        }
    }
}