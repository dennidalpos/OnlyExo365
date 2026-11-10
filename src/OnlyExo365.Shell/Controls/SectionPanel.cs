using System.Windows;
using System.Windows.Controls;

namespace OnlyExo365.Shell.Controls;

public class SectionPanel : HeaderedContentControl
{
    static SectionPanel()
    {
        DefaultStyleKeyProperty.OverrideMetadata(
            typeof(SectionPanel),
            new FrameworkPropertyMetadata(typeof(SectionPanel)));
    }
}

