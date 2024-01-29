using System.Diagnostics;

class Program
{
    static void Main()
    {
        // Menuweb
        string Menuweb = "http://menuweb.centermastersul.com.br";

        // Abrir a URL no navegador padrão
        Process.Start(new ProcessStartInfo
        {
            FileName = Menuweb,
            UseShellExecute = true
        });
    }
}
