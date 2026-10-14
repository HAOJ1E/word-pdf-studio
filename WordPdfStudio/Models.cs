using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.IO;
namespace WordPdfStudio;
public sealed class QueueItem : INotifyPropertyChanged
{
    public string SourcePath { get; set; } = "";
    public string Name => Path.GetFileName(SourcePath);
    public string Type => Path.GetExtension(SourcePath).TrimStart('.').ToUpperInvariant();
    public string Size { get; set; } = "";
    private string _status = "等待转换", _detail = "", _output = "";
    public string Status { get => _status; set { _status=value; OnChanged(); OnChanged(nameof(StatusColor)); } }
    public string Detail { get => _detail; set { _detail=value; OnChanged(); } }
    public string OutputPath { get => _output; set { _output=value; OnChanged(); OnChanged(nameof(CanOpen)); } }
    public bool CanOpen => !string.IsNullOrEmpty(OutputPath) && File.Exists(OutputPath);
    public string StatusColor => Status switch { "转换成功"=>"#248875", "转换失败"=>"#C25464", "转换中"=>"#6855E8", _=>"#8390A8" };
    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnChanged([CallerMemberName]string? name=null) => PropertyChanged?.Invoke(this,new PropertyChangedEventArgs(name));
}
public sealed record HistoryItem(DateTime Time,string Action,string Source,string Output,bool Success,string Detail)
{
    public string TimeLabel => Time.ToString("MM-dd HH:mm");
    public string FileName => Path.GetFileName(Output.Length>0 ? Output : Source);
    public string Status => Success ? "成功" : "失败";
}
public sealed class AppSettings
{
    public string OutputFolder { get; set; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),"文页PDF");
    public int QualityIndex { get; set; } = 0;
    public bool SourceFolder { get; set; }
    public bool OpenAfter { get; set; } = true;
}
public static class LocalStore
{
    public static string Root => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"WordPdfStudio");
    public static T Read<T>(string name,T fallback) { try { return JsonSerializer.Deserialize<T>(File.ReadAllText(Path.Combine(Root,name))) ?? fallback; } catch { return fallback; } }
    public static void Write<T>(string name,T data)
    {
        Directory.CreateDirectory(Root);
        string path=Path.Combine(Root,name), temp=path+"."+Guid.NewGuid().ToString("N")+".tmp";
        try { File.WriteAllText(temp,JsonSerializer.Serialize(data,new JsonSerializerOptions { WriteIndented=true })); File.Move(temp,path,true); }
        finally { if(File.Exists(temp)) File.Delete(temp); }
    }
    public static string UniquePath(string desired)
    {
        if(!File.Exists(desired) && !Directory.Exists(desired)) return desired;
        string folder=Path.GetDirectoryName(desired)!, name=Path.GetFileNameWithoutExtension(desired), ext=Path.GetExtension(desired);
        for(int n=2;n<100000;n++) { var candidate=Path.Combine(folder,$"{name}_{n}{ext}"); if(!File.Exists(candidate) && !Directory.Exists(candidate)) return candidate; }
        throw new IOException("同名文件过多，请换一个输出目录。");
    }
}
