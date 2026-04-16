using ExchangeAdmin.Application.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class ResourcesViewModel : ResourcesPageViewModel
{
    public ResourcesViewModel(IResourcesWorkerService workerService, ShellViewModel shellViewModel)
        : base(workerService, shellViewModel)
    {
    }
}
