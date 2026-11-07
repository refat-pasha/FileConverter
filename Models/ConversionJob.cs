using System.ComponentModel;
using System.IO;
using FileConverter.Services;

namespace FileConverter.Models;

public enum JobStatus { Pending, Running, Done, Failed }

public class ConversionJob : INotifyPropertyChanged
{
    public string SourcePath { get; }
    public string FileName => Path.GetFileName(SourcePath);
    public string Extension => Path.GetExtension(SourcePath).ToLowerInvariant();

    public TargetFormat[] AvailableTargets { get; }

    private TargetFormat _targetFormat;
    public TargetFormat TargetFormat
    {
        get => _targetFormat;
        set { if (Set(ref _targetFormat, value)) ResetStatus(); }
    }

    private JobStatus _status;
    public JobStatus Status
    {
        get => _status;
        set
        {
            if (Set(ref _status, value))
            {
                OnPropertyChanged(nameof(StatusText));
                OnPropertyChanged(nameof(Detail));
            }
        }
    }

    public string StatusText => Status switch
    {
        JobStatus.Pending => "Pending",
        JobStatus.Running => "Converting…",
        JobStatus.Done => "Done",
        JobStatus.Failed => "Failed",
        _ => ""
    };

    private string? _outputPath;
    public string? OutputPath { get => _outputPath; set { if (Set(ref _outputPath, value)) OnPropertyChanged(nameof(Detail)); } }

    private string? _error;
    public string? Error { get => _error; set { if (Set(ref _error, value)) OnPropertyChanged(nameof(Detail)); } }

    /// <summary>Short status detail shown in the grid: output name or error summary.</summary>
    public string? Detail => Status switch
    {
        JobStatus.Done => OutputPath is null ? null : Path.GetFileName(OutputPath),
        JobStatus.Failed => Error,
        _ => null
    };

    public ConversionJob(string sourcePath, TargetFormat defaultTarget)
    {
        SourcePath = Path.GetFullPath(sourcePath);
        AvailableTargets = ConversionCatalog.TargetsFor(Extension);
        _targetFormat = defaultTarget;
    }

    private void ResetStatus()
    {
        if (Status is JobStatus.Done or JobStatus.Failed)
        {
            Status = JobStatus.Pending;
            OutputPath = null;
            Error = null;
            OnPropertyChanged(nameof(Detail));
        }
    }

    public void RefreshDetail() => OnPropertyChanged(nameof(Detail));

    public event PropertyChangedEventHandler? PropertyChanged;

    private bool Set<T>(ref T field, T value, [System.Runtime.CompilerServices.CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        return true;
    }

    private void OnPropertyChanged(string name) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
