using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using WordPdfStudio.Services;
namespace WordPdfStudio;
public partial class MainWindow : Window
{
    private readonly ObservableCollection<QueueItem> _queue = new();
    private readonly ObservableCollection<QueueItem> _toolFiles = new();
    private readonly ObservableCollection<HistoryItem> _history;
    private readonly EngineManager _engine = new();
    private readonly AppSettings _settings;
    private CancellationTokenSource? _cancel;
    private bool _busy, _closeAfterCancel, _initialized;
    private string _page = "convert", _tool = "merge";
    public MainWindow()
    {
        InitializeComponent();
        _settings=LocalStore.Read("settings.json",new AppSettings());
        _history=new(LocalStore.Read("history.json",new List<HistoryItem>()).Take(200));
        QueueGrid.ItemsSource=_queue; ToolFilesList.ItemsSource=_toolFiles; HistoryGrid.ItemsSource=_history;
        OutputFolderBox.Text=_settings.OutputFolder; QualityBox.SelectedIndex=Math.Clamp(_settings.QualityIndex,0,2);
        SourceFolderCheck.IsChecked=_settings.SourceFolder; OpenAfterCheck.IsChecked=_settings.OpenAfter;
        _initialized=true; SourceFolderChanged(this,new RoutedEventArgs()); Navigate("convert"); SetTool("merge"); UpdateCounts();
    }
    private static Brush Brush(string color) => new SolidColorBrush((Color)ColorConverter.ConvertFromString(color));
    private static void ShellOpen(string path)
    {
        if(!File.Exists(path) && !Directory.Exists(path)) throw new IOException("文件已移动或删除。请检查原输出位置。");
        Process.Start(new ProcessStartInfo(path) { UseShellExecute=true });
    }
    private void Notice(string text) => FooterStatus.Text=text;
    private void ShowError(Exception error) { Notice(error.Message); MessageBox.Show(this,error.Message,"文页 PDF",MessageBoxButton.OK,MessageBoxImage.Warning); }
    public void Navigate(string page)
    {
        _page=page is "tools" or "history" or "about" ? page : "convert";
        ConvertPanel.Visibility=_page=="convert" ? Visibility.Visible : Visibility.Collapsed;
        ToolsPanel.Visibility=_page=="tools" ? Visibility.Visible : Visibility.Collapsed;
        HistoryPanel.Visibility=_page=="history" ? Visibility.Visible : Visibility.Collapsed;
        AboutPanel.Visibility=_page=="about" ? Visibility.Visible : Visibility.Collapsed;
        var names=new Dictionary<string,(string,string)>{["convert"]=("Word 转 PDF","从一份文档，到一整个文件夹，轻松完成转换。"),["tools"]=("PDF 工具箱","整理、保护与分享，常用工具一次备齐。"),["history"]=("处理记录","每次处理都有迹可循，随时找到输出文件。"),["about"]=("使用指南","了解运行方式，快速上手文页 PDF。")};
        (PageTitle.Text,PageSubtitle.Text)=names[_page];
        foreach(var nav in new[]{NavConvert,NavTools,NavHistory,NavAbout})
        { bool active=(string)nav.Tag==_page; nav.Background=Brush(active?"#EEEAFE":"#FFFFFF"); nav.Foreground=Brush(active?"#6855E8":"#78829B"); nav.BorderThickness=new Thickness(0); }
    }
    private void NavClick(object sender,RoutedEventArgs e) => Navigate((string)((Button)sender).Tag);
    public void AddPaths(IEnumerable<string> paths)
    {
        if(_busy) { Notice("正在处理，请完成或取消后再添加文件。"); return; }
        var target=_page=="tools" ? _toolFiles : _queue;
        int added=0, skipped=0;
        foreach(string raw in paths)
        {
            try
            {
                string path=Path.GetFullPath(raw);
                if(Directory.Exists(path))
                {
                    if(target==_toolFiles) { skipped++; continue; }
                    foreach(string file in EnumerateDocuments(path)) if(AddOne(file,target)) added++; else skipped++;
                }
                else if(AddOne(path,target)) added++; else skipped++;
            }
            catch(Exception ex) when(ex is IOException or UnauthorizedAccessException or ArgumentException) { skipped++; }
        }
        UpdateCounts(); Notice($"已添加 {added} 个文件"+(skipped>0?$"，跳过 {skipped} 个不支持、重复或无法读取的文件。":"。"));
    }
    private IEnumerable<string> EnumerateDocuments(string root)
    {
        var options=new EnumerationOptions{RecurseSubdirectories=true,IgnoreInaccessible=true,AttributesToSkip=FileAttributes.ReparsePoint};
        return Directory.EnumerateFiles(root,"*",options).Where(ConversionService.IsSupportedFile).Where(p=>!Path.GetFileName(p).StartsWith("~$",StringComparison.Ordinal)).OrderBy(p=>p,StringComparer.OrdinalIgnoreCase);
    }
    private bool AddOne(string path,ObservableCollection<QueueItem> target)
    {
        if(!File.Exists(path) || target.Any(q=>string.Equals(q.SourcePath,path,StringComparison.OrdinalIgnoreCase))) return false;
        bool supported=target==_queue ? ConversionService.IsSupportedFile(path) : _tool=="images" ? new[]{".png",".jpg",".jpeg",".bmp",".tif",".tiff"}.Contains(Path.GetExtension(path).ToLowerInvariant()) : Path.GetExtension(path).Equals(".pdf",StringComparison.OrdinalIgnoreCase);
        if(!supported || Path.GetFileName(path).StartsWith("~$",StringComparison.Ordinal)) return false;
        long length=new FileInfo(path).Length;
        target.Add(new QueueItem{SourcePath=path,Size=length>=1048576?$"{length/1048576d:0.0} MB":$"{Math.Max(1,length/1024d):0} KB"}); return true;
    }
    private void WindowDragOver(object sender,DragEventArgs e) { e.Effects=!_busy && e.Data.GetDataPresent(DataFormats.FileDrop)?DragDropEffects.Copy:DragDropEffects.None; e.Handled=true; }
    private void WindowDrop(object sender,DragEventArgs e) { if(e.Data.GetData(DataFormats.FileDrop) is string[] paths) AddPaths(paths); e.Handled=true; }
    private void AddFilesClick(object sender,RoutedEventArgs e)
    { var dialog=new OpenFileDialog{Filter="Word 与文本文档|*.docx;*.doc;*.rtf;*.odt",Multiselect=true,Title="添加文档"}; if(dialog.ShowDialog(this)==true) AddPaths(dialog.FileNames); }
    private void AddFolderClick(object sender,RoutedEventArgs e)
    { var dialog=new OpenFolderDialog{Title="添加整个文件夹（包括子文件夹）"}; if(dialog.ShowDialog(this)==true) AddPaths(new[]{dialog.FolderName}); }
    private void ChooseOutputClick(object sender,RoutedEventArgs e)
    { var dialog=new OpenFolderDialog{Title="选择输出目录"}; if(Directory.Exists(OutputFolderBox.Text)) dialog.InitialDirectory=OutputFolderBox.Text; if(dialog.ShowDialog(this)==true) OutputFolderBox.Text=dialog.FolderName; }
    private void SourceFolderChanged(object sender,RoutedEventArgs e) { if(_initialized) OutputSettings.IsEnabled=!_busy && SourceFolderCheck.IsChecked!=true; }
    private void PdfACheckChanged(object sender,RoutedEventArgs e)
    { if(MergeCheck==null)return; bool archival=PdfACheck.IsChecked==true; if(archival) MergeCheck.IsChecked=false; MergeCheck.IsEnabled=!archival; }
    private void RemoveClick(object sender,RoutedEventArgs e) { foreach(var q in QueueGrid.SelectedItems.Cast<QueueItem>().ToArray()) _queue.Remove(q); UpdateCounts(); }
    private void ClearDoneClick(object sender,RoutedEventArgs e) { foreach(var q in _queue.Where(q=>q.Status=="转换成功").ToArray()) _queue.Remove(q); UpdateCounts(); }
    private void MoveUpClick(object sender,RoutedEventArgs e) => MoveItem(_queue,QueueGrid.SelectedItem as QueueItem,-1);
    private void MoveDownClick(object sender,RoutedEventArgs e) => MoveItem(_queue,QueueGrid.SelectedItem as QueueItem,1);
    private static void MoveItem(ObservableCollection<QueueItem> list,QueueItem? item,int direction)
    { if(item==null)return; int current=list.IndexOf(item), next=current+direction; if(next>=0 && next<list.Count) list.Move(current,next); }
    private void QueueSelectionChanged(object sender,SelectionChangedEventArgs e)
    { if(QueueGrid.SelectedItem is QueueItem q) Notice(q.Detail.Length>0?q.Detail:q.SourcePath); }
    private void QueueDoubleClick(object sender,MouseButtonEventArgs e)
    { if(QueueGrid.SelectedItem is not QueueItem q)return; if(q.CanOpen) { try{ShellOpen(q.OutputPath);}catch(Exception ex){ShowError(ex);} } else if(q.Detail.Length>0) MessageBox.Show(this,q.Detail,q.Name); }
    private void UpdateCounts()
    {
        QueuedCount.Text=_queue.Count(q=>q.Status is "等待转换" or "已取消").ToString();
        DoneCount.Text=_queue.Count(q=>q.Status=="转换成功").ToString(); ErrorCount.Text=_queue.Count(q=>q.Status=="转换失败").ToString();
        EmptyHint.Visibility=_queue.Count==0?Visibility.Visible:Visibility.Collapsed;
        StartButton.IsEnabled=!_busy && _queue.Any(q=>q.Status!="转换成功");
        StartButton.Content=_queue.Any(q=>q.Status=="转换失败")?"转换 / 重试  →":"开始转换  →";
    }
    private void SetBusy(bool value,bool cancellable=false)
    {
        _busy=value; QueueAddActions.IsEnabled=!value; QueueEditActions.IsEnabled=!value;
        ConversionSettings.IsEnabled=!value; OutputFlags.IsEnabled=!value; OutputSettings.IsEnabled=!value && SourceFolderCheck.IsChecked!=true;
        ToolChoices.IsEnabled=!value; ToolFileActions.IsEnabled=!value; ToolOptions.IsEnabled=!value; RunToolButton.IsEnabled=!value;
        CancelButton.IsEnabled=value && cancellable; UpdateCounts();
    }
    private void SaveSettings()
    {
        _settings.OutputFolder=OutputFolderBox.Text; _settings.QualityIndex=QualityBox.SelectedIndex; _settings.SourceFolder=SourceFolderCheck.IsChecked==true; _settings.OpenAfter=OpenAfterCheck.IsChecked==true;
        try { LocalStore.Write("settings.json",_settings); } catch(Exception ex) { Notice("无法保存设置："+ex.Message); }
    }
    private void AddHistory(string action,string source,string output,bool success,string detail)
    {
        _history.Insert(0,new HistoryItem(DateTime.Now,action,source,output,success,detail));
        while(_history.Count>200)_history.RemoveAt(_history.Count-1);
        try{LocalStore.Write("history.json",_history);}catch(Exception ex){Notice("结果已保存，但无法保存记录："+ex.Message);}
    }
    private async void StartConversionClick(object sender,RoutedEventArgs e)
    {
        if(_busy)return;
        var jobs=_queue.Where(q=>q.Status!="转换成功").ToArray(); if(jobs.Length==0)return;
        ConversionOptions options;
        string outputFolder;
        bool sourceFolder=SourceFolderCheck.IsChecked==true, merge=MergeCheck.IsChecked==true, open=OpenAfterCheck.IsChecked==true;
        try
        {
            if(string.IsNullOrWhiteSpace(OutputFolderBox.Text)) throw new ArgumentException("请选择输出目录。");
            outputFolder=Path.GetFullPath(OutputFolderBox.Text.Trim());
            if(!sourceFolder || merge) Directory.CreateDirectory(outputFolder);
            options=new ConversionOptions{Quality=QualityBox.SelectedIndex switch{1=>82,2=>60,_=>95},PdfA=PdfACheck.IsChecked==true,PageRange=ConversionService.NormalizePageRange(ConvertRangeBox.Text)};
        }catch(Exception ex){ShowError(ex);return;}
        SaveSettings(); _cancel=new CancellationTokenSource(); SetBusy(true,true);
        int done=0,fail=0; var succeeded=new List<string>(); string? lastOutput=null;
        try
        {
            var converter=new ConversionService(_engine);
            for(int index=0;index<jobs.Length;index++)
            {
                _cancel.Token.ThrowIfCancellationRequested();
                var item=jobs[index]; item.Status="转换中"; item.Detail="正在准备转换…"; QueueGrid.ScrollIntoView(item);
                try
                {
                    string folder=sourceFolder?Path.GetDirectoryName(item.SourcePath)!:outputFolder;
                    string output=LocalStore.UniquePath(Path.Combine(folder,Path.GetFileNameWithoutExtension(item.SourcePath)+".pdf"));
                    var progress=new Progress<EngineProgress>(p=>{BatchStatus.Text=$"{index+1}/{jobs.Length} · {p.Message}"; BatchProgress.Value=(index+Math.Clamp(p.Fraction,0,1))*100/jobs.Length;});
                    await converter.ConvertAsync(item.SourcePath,output,options,progress,_cancel.Token);
                    item.OutputPath=output; item.Status="转换成功"; item.Detail=output; succeeded.Add(output); lastOutput=output; done++;
                    AddHistory("Word 转 PDF",item.SourcePath,output,true,"转换完成。");
                }
                catch(OperationCanceledException){item.Status="已取消"; item.Detail="已取消，点击开始转换可重新处理。"; throw;}
                catch(Exception ex){item.Status="转换失败";item.Detail=ex.Message;fail++;AddHistory("Word 转 PDF",item.SourcePath,"",false,ex.Message);}
                BatchProgress.Value=(index+1)*100d/jobs.Length; UpdateCounts();
            }
            if(merge && succeeded.Count>=2)
            {
                _cancel.Token.ThrowIfCancellationRequested();
                string merged=LocalStore.UniquePath(Path.Combine(outputFolder,"合并结果.pdf")); BatchStatus.Text="正在合并本次成功转换的 PDF…";
                await Task.Run(()=>PdfService.Merge(succeeded,merged),_cancel.Token); lastOutput=merged; AddHistory("转换后合并",string.Join("\n",succeeded),merged,true,$"合并本次 {succeeded.Count} 个成功文件。");
            }
            string mergeNote=merge && succeeded.Count<2?"；成功文件不足 2 个，未生成合并文件":"";
            BatchStatus.Text=$"处理完成 · 成功 {done} 个，失败 {fail} 个{mergeNote}"; Notice(fail>0?"失败文件可双击查看原因，再点击转换 / 重试。":"转换已完成，双击队列打开 PDF。");
            if(open && lastOutput!=null) { try{ShellOpen(Path.GetDirectoryName(lastOutput)!);}catch(Exception ex){Notice("转换完成，打开目录失败："+ex.Message);} }
        }
        catch(OperationCanceledException)
        {
            foreach(var job in jobs.Where(q=>q.Status is "等待转换" or "转换中"))job.Status="已取消";
            BatchStatus.Text=$"已取消 · 已完成 {done} 个，失败 {fail} 个"; Notice("已完成的文件保留；再次开始可继续处理未完成文件。");
        }
        catch(Exception ex){BatchStatus.Text="处理结束："+ex.Message;ShowError(ex);}
        finally { _cancel.Dispose(); _cancel=null; SetBusy(false); if(_closeAfterCancel) Close(); }
    }
    private void CancelClick(object sender,RoutedEventArgs e) { _cancel?.Cancel(); CancelButton.IsEnabled=false; BatchStatus.Text="正在取消，请稍候…"; }
    private void SetTool(string tool)
    {
        bool switchedImage=(_tool=="images")!=(tool=="images"); _tool=tool; if(switchedImage)_toolFiles.Clear();
        var info=new Dictionary<string,(string,string)>{["merge"]=("合并 PDF","添加两份或更多 PDF，按列表顺序拼接。"),["extract"]=("提取页面","添加一份 PDF，按指定顺序提取页面。"),["split"]=("逐页拆分","添加一份 PDF，每一页生成一个独立文件。"),["watermark"]=("添加水印","添加一份 PDF，为所有页面添加半透明文字水印。"),["encrypt"]=("密码保护","添加一份 PDF，设置打开密码。"),["rotate"]=("旋转页面","添加一份 PDF，顺时针旋转所有页面。"),["images"]=("图片转 PDF","支持 PNG / JPG / BMP / TIFF；每张图片生成一页。")};
        (ToolTitle.Text,ToolHint.Text)=info[tool];
        RangeOption.Visibility=tool=="extract"?Visibility.Visible:Visibility.Collapsed;
        WatermarkOption.Visibility=tool=="watermark"?Visibility.Visible:Visibility.Collapsed;
        PasswordOption.Visibility=tool=="encrypt"?Visibility.Visible:Visibility.Collapsed;
        RotateOption.Visibility=tool=="rotate"?Visibility.Visible:Visibility.Collapsed;
        InputPasswordOption.Visibility=tool!="images"?Visibility.Visible:Visibility.Collapsed;
        foreach(var child in ToolChoices.Children.OfType<Button>()) { bool active=(string)child.Tag==tool; child.Background=Brush(active?"#EEEAFE":"#FFFFFF"); child.Foreground=Brush(active?"#6855E8":"#78829B"); }
        ToolStatus.Text="添加文件后，点击右侧按钮选择保存位置。";
        ToolNote.Text=tool=="images"?"按列表顺序生成 A4 页面，自动选择横竖方向；多帧 TIFF 仅使用首帧。":"生成新文件，原文件保留。除密码保护功能外，新文件不保留原打开密码；书签、交互表单和数字签名可能不保留。";
        if(tool!="encrypt"){NewPasswordBox.Clear();ConfirmPasswordBox.Clear();}
    }
    private void ToolClick(object sender,RoutedEventArgs e) => SetTool((string)((Button)sender).Tag);
    private void AddToolFilesClick(object sender,RoutedEventArgs e)
    { var dialog=new OpenFileDialog{Filter=_tool=="images"?"图片|*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff":"PDF 文件|*.pdf",Multiselect=_tool is "merge" or "images",Title="添加处理文件"};if(dialog.ShowDialog(this)==true)AddPaths(dialog.FileNames); }
    private void RemoveToolFileClick(object sender,RoutedEventArgs e) { if(ToolFilesList.SelectedItem is QueueItem q)_toolFiles.Remove(q); }
    private void ClearToolFilesClick(object sender,RoutedEventArgs e)=>_toolFiles.Clear();
    private void MoveToolUpClick(object sender,RoutedEventArgs e)=>MoveItem(_toolFiles,ToolFilesList.SelectedItem as QueueItem,-1);
    private void MoveToolDownClick(object sender,RoutedEventArgs e)=>MoveItem(_toolFiles,ToolFilesList.SelectedItem as QueueItem,1);
    private async void RunToolClick(object sender,RoutedEventArgs e)
    {
        if(_busy)return; var inputs=_toolFiles.Select(f=>f.SourcePath).ToArray(); string tool=_tool;
        if(inputs.Length==0){Notice("请先添加文件。");return;}
        if(tool=="merge" && inputs.Length<2){Notice("合并需要至少两份 PDF。");return;}
        if(tool is not ("merge" or "images") && inputs.Length!=1){Notice("此工具每次处理一份 PDF，请移除多余文件。");return;}
        string range=ToolRangeBox.Text.Trim(), watermark=WatermarkBox.Text.Trim(), password=NewPasswordBox.Password, inputPassword=InputPasswordBox.Password;
        if(tool=="encrypt" && (password.Length<1 || password!=ConfirmPasswordBox.Password)){Notice("请设置密码，并确保两次输入一致。");return;}
        if(tool=="watermark" && watermark.Length==0){Notice("请输入水印文字。");return;}
        int rotation=(RotationBox.SelectedIndex+1)*90; string output;
        if(tool=="split") { var dialog=new OpenFolderDialog{Title="选择拆分结果目录"}; if(dialog.ShowDialog(this)!=true)return; output=dialog.FolderName; }
        else
        {
            var dialog=new SaveFileDialog{Filter="PDF 文件|*.pdf",DefaultExt=".pdf",AddExtension=true,OverwritePrompt=false,FileName=tool is "merge" or "images"?ToolTitle.Text+"结果.pdf":Path.GetFileNameWithoutExtension(inputs[0])+"_"+ToolTitle.Text+".pdf"};
            if(Directory.Exists(OutputFolderBox.Text))dialog.InitialDirectory=OutputFolderBox.Text;
            if(dialog.ShowDialog(this)!=true)return; output=LocalStore.UniquePath(dialog.FileName);
        }
        SetBusy(true); ToolStatus.Text="正在处理，请稍候…";
        try
        {
            IReadOnlyList<string>? results=null;
            await Task.Run(()=>
            {
                switch(tool)
                {
                    case "merge":PdfService.Merge(inputs,output,inputPassword);break;
                    case "extract":PdfService.Extract(inputs[0],output,range,inputPassword);break;
                    case "split":results=PdfService.Split(inputs[0],output,inputPassword);break;
                    case "watermark":PdfService.Watermark(inputs[0],output,watermark,inputPassword);break;
                    case "encrypt":PdfService.Encrypt(inputs[0],output,password,inputPassword);break;
                    case "rotate":PdfService.Rotate(inputs[0],output,rotation,inputPassword);break;
                    case "images":PdfService.ImagesToPdf(inputs,output);break;
                }
            });
            ToolStatus.Text=tool=="split"?$"完成 · 已拆分 {results?.Count} 页，保存到 {output}":$"完成 · {output}";
            AddHistory(ToolTitle.Text,string.Join("\n",inputs),output,true,ToolStatus.Text);
            Notice("处理完成，结果可从处理记录中打开。"); NewPasswordBox.Clear();ConfirmPasswordBox.Clear();InputPasswordBox.Clear();
            if(OpenAfterCheck.IsChecked==true)try{ShellOpen(tool=="split"?output:Path.GetDirectoryName(output)!);}catch(Exception ex){Notice("处理完成，打开目录失败："+ex.Message);}
        }
        catch(Exception ex){ToolStatus.Text="处理失败："+ex.Message;AddHistory(ToolTitle.Text,string.Join("\n",inputs),"",false,ex.Message);ShowError(ex);}
        finally{SetBusy(false);if(_closeAfterCancel)Close();}
    }
    private void ClearHistoryClick(object sender,RoutedEventArgs e)
    { _history.Clear();try{LocalStore.Write("history.json",_history);Notice("记录已清空，输出文件保留。");}catch(Exception ex){ShowError(ex);} }
    private void OpenHistoryClick(object sender,RoutedEventArgs e)
    { if(HistoryGrid.SelectedItem is HistoryItem h && h.Success)try{ShellOpen(h.Output);}catch(Exception ex){ShowError(ex);} }
    private void HistoryDoubleClick(object sender,MouseButtonEventArgs e)=>OpenHistoryClick(sender,new RoutedEventArgs());
    private void HistoryDetailClick(object sender,RoutedEventArgs e)
    { if(HistoryGrid.SelectedItem is HistoryItem h)MessageBox.Show(this,$"{h.Time:yyyy-MM-dd HH:mm:ss}\n{h.Action} · {h.Status}\n\n来源：\n{h.Source}\n\n结果：\n{h.Output}\n\n{h.Detail}","处理详情"); }
    private void OpenDataClick(object sender,RoutedEventArgs e) { try{Directory.CreateDirectory(LocalStore.Root);ShellOpen(LocalStore.Root);}catch(Exception ex){ShowError(ex);} }
    private void WindowClosing(object? sender,CancelEventArgs e)
    {
        if(_busy){e.Cancel=true;_closeAfterCancel=true;_cancel?.Cancel();Notice("等待当前处理结束后关闭窗口…");}
        else if(!App.IsScreenshotMode)SaveSettings();
    }
}
