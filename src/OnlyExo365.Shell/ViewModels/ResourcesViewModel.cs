using OnlyExo365.Shell.Services;

namespace OnlyExo365.Shell.ViewModels;

public sealed class ResourcesViewModel : ResourcesPageViewModel
{
    public ResourcesViewModel(IResourcesWorkerService workerService, ShellViewModel shellViewModel)
        : base(workerService, shellViewModel)
    {
    }
}

