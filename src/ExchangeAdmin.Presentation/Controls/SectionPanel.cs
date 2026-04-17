using System.Windows;
using System.Windows.Controls;

namespace ExchangeAdmin.Presentation.Controls;

public class SectionPanel : HeaderedContentControl
{
    static SectionPanel()
    {
        DefaultStyleKeyProperty.OverrideMetadata(
            typeof(SectionPanel),
            new FrameworkPropertyMetadata(typeof(SectionPanel)));
    }
}
