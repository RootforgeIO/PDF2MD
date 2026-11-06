using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        string root = AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string pythonw = Path.Combine(root, "runtime", "pythonw.exe");
        string python = Path.Combine(root, "runtime", "python.exe");
        string app = Path.Combine(root, "app", "pdf2md_ui.py");
#if INSTALLED
        string logs = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PDF2MD", "logs");
#else
        string logs = Path.Combine(root, "logs");
        string models = Path.Combine(root, "models");
        Directory.CreateDirectory(models);
#endif

        string interpreter = File.Exists(pythonw) ? pythonw : python;
        if (!File.Exists(interpreter))
        {
            MessageBox.Show("Portable runtime not found:\n" + interpreter, "PDF2MD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 2;
        }

        if (!File.Exists(app))
        {
            MessageBox.Show("Application entrypoint not found:\n" + app, "PDF2MD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 3;
        }

        string runtime = Path.Combine(root, "runtime");
        string scripts = Path.Combine(runtime, "Scripts");

        var psi = new ProcessStartInfo();
        psi.FileName = interpreter;
        psi.Arguments = Quote(app);
        psi.WorkingDirectory = Path.Combine(root, "app");
        psi.UseShellExecute = false;

        string oldPath = Environment.GetEnvironmentVariable("PATH") ?? "";
        psi.EnvironmentVariables["PATH"] = runtime + ";" + scripts + ";" + oldPath;
#if INSTALLED
        psi.Arguments = "-E -s -B " + Quote(app);
        psi.EnvironmentVariables.Remove("PDF2MD_PORTABLE_ROOT");
#else
        string sitePackages = Path.Combine(runtime, "Lib", "site-packages");
        psi.EnvironmentVariables["PYTHONPATH"] = Path.Combine(root, "app") + ";" + sitePackages;
        psi.EnvironmentVariables["PDF2MD_PORTABLE_ROOT"] = root;
        psi.EnvironmentVariables["MODEL_CACHE_DIR"] = Path.Combine(models, "datalab", "models");
        psi.EnvironmentVariables["HF_HOME"] = Path.Combine(models, "huggingface");
        psi.EnvironmentVariables["HUGGINGFACE_HUB_CACHE"] = Path.Combine(models, "huggingface", "hub");
        psi.EnvironmentVariables["TRANSFORMERS_CACHE"] = Path.Combine(models, "huggingface", "transformers");
        psi.EnvironmentVariables["TORCH_HOME"] = Path.Combine(models, "torch");
        psi.EnvironmentVariables["XDG_CACHE_HOME"] = Path.Combine(models, "cache");
#endif

        try
        {
            Process.Start(psi);
            return 0;
        }
        catch (Exception ex)
        {
            try
            {
                Directory.CreateDirectory(logs);
                File.WriteAllText(Path.Combine(logs, "launcher-error.txt"), ex.ToString());
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            MessageBox.Show(ex.Message, "PDF2MD launcher error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 4;
        }
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }
}
