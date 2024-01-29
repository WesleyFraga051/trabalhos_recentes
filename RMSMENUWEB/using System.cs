using System.Diagnostics;

class Program
{
    static void Main()
    {
        // URL do YouTube
        string youtubeUrl = "https://www.youtube.com";

        // Abrir o YouTube no navegador padrão
        Process.Start(new ProcessStartInfo(youtubeUrl) { UseShellExecute = true });
    }
}
