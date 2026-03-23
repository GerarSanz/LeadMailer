using System.IO;
using LeadMailer.Models;
using OfficeOpenXml;

namespace LeadMailer.Services;

/// <summary>Lee leads desde un archivo Excel (.xlsx / .xls).</summary>
public class ExcelService
{
    private static bool _licenseConfigured;

    private static void EnsureEpplusLicense()
    {
        if (_licenseConfigured) return;

        ExcelPackage.License.SetNonCommercialPersonal("LeadMailer");

        _licenseConfigured = true;
    }

    public List<Lead> ReadLeads(string filePath)
    {
        EnsureEpplusLicense();
        var leads = new List<Lead>();

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"No se encontró el archivo: {filePath}");

        using var package = new ExcelPackage(new FileInfo(filePath));
        var ws = package.Workbook.Worksheets[0];
        if (ws?.Dimension == null) return leads;

        // ── Mapear columnas por cabecera ──────────────────────────────────────
        var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = 1; c <= ws.Dimension.End.Column; c++)
            colMap[ws.Cells[1, c].Text.Trim()] = c;

        int Col(params string[] names)
        {
            foreach (var n in names)
                if (colMap.TryGetValue(n, out var col)) return col;
            return -1;
        }

        int cFecha    = Col("  ", "Fecha", "Timestamp");
        int cCurso    = Col("Curso");
        int cPlat     = Col("Plataforma");
        int cSit      = Col("Situación Laboral", "Situacion Laboral");
        int cNivel    = Col("Sector Laboral / Nivel Estudios", "Nivel Estudios");
        int cNombre   = Col("Nombre");
        int cEmail    = Col("Email", "Correo", "E-mail");
        int cTel      = Col("Teléfono", "Telefono");
        int cProv     = Col("Provincia");
        int cPub      = Col("Aceptación Publicidad", "Aceptacion Publicidad");
        int cObs      = Col("OBSERVACIONES", "Observaciones");
        int cContact  = Col("CONTACTO. PERSONA / FECHA / MEDIO / RESULTADO", "Contacto");
        int cInsc     = Col("INSCRIPCIÓN (SI/NO)", "Inscripción", "Inscripcion");

        string Cell(int row, int col) => col > 0 ? ws.Cells[row, col].Text.Trim() : "";
        
        // Lectura especial para teléfono: asegurar que se lee como texto puro
        string CellAsText(int row, int col)
        {
            if (col <= 0) return "";
            var cell = ws.Cells[row, col];
            // Intentar leer como string primero
            if (cell.Value is string str)
                return str.Trim();
            // Si es número, convertir a string sin formato científico
            if (cell.Value is double d)
                return d.ToString("0");
            if (cell.Value is long l)
                return l.ToString();
            if (cell.Value is int i)
                return i.ToString();
            // Fallback al texto normal
            return cell.Text.Trim();
        }

        // ── Leer filas ────────────────────────────────────────────────────────
        for (int row = 2; row <= ws.Dimension.End.Row; row++)
        {
            var email = Cell(row, cEmail);
            var curso = Cell(row, cCurso);
            if (string.IsNullOrWhiteSpace(email) && string.IsNullOrWhiteSpace(curso)) continue;

            var pubText = Cell(row, cPub).ToLowerInvariant();
            bool acepta = pubText is "true" or "1" or "sí" or "si" or "yes";

            leads.Add(new Lead
            {
                SourceRow        = row,
                FechaRegistro    = Cell(row, cFecha),
                Curso            = curso,
                Plataforma       = Cell(row, cPlat),
                SituacionLaboral = Cell(row, cSit),
                NivelEstudios    = Cell(row, cNivel),
                Nombre           = Cell(row, cNombre),
                Email            = email,
                Telefono         = CellAsText(row, cTel),  // Usar CellAsText para evitar formato científico
                Provincia        = Cell(row, cProv),
                AceptaPublicidad = acepta,
                Observaciones    = Cell(row, cObs),
                Contacto         = Cell(row, cContact),
                Inscripcion      = Cell(row, cInsc)
            });
        }

        return leads;
    }
}
