---
name: create-floating-window
description: Create floating WPF windows, popups, context menus, floating toolbars, modeless dialogs, dark-themed panels, window styling, drag-to-move, DPI-aware positioning, transparent windows
---

# Create Floating Window Skill

This skill provides guidelines and templates for creating new floating WPF windows in the oPenEfficiency project.

## Requirements for Floating Windows
When creating a new floating window (such as a context menu, settings panel, or tool palette), ALWAYS adhere to the following properties to ensure consistency and correct behavior within the PowerPoint Add-in context.

### 1. Window Properties
In the `<Window>` tag, include these attributes:
```xml
WindowStyle="None"
AllowsTransparency="True"
Background="Transparent"
Topmost="True" 
ShowInTaskbar="False"
WindowStartupLocation="CenterScreen" 
MouseDown="Window_MouseDown"
FontFamily="Segoe UI Variable Small, Segoe UI"
```
*Note: If implementing a context menu that appears at the cursor, remove `WindowStartupLocation` and set `.Left` and `.Top` in the code-behind using `System.Windows.Forms.Cursor.Position`.*

### 2. Root Element Styling
Always use a curved, dark-themed `Border` with a drop shadow as the root element inside the window to create a floating panel effect.

```xml
<Window.Resources>
    <ResourceDictionary>
        <ResourceDictionary.MergedDictionaries>
            <ResourceDictionary Source="UI/Styles/Controls.xaml"/>
        </ResourceDictionary.MergedDictionaries>
        
        <!-- Add BooleanToVisibilityConverter if needed for tab visibility -->
        <BooleanToVisibilityConverter x:Key="BooleanToVisibilityConverter"/>
        
        <!-- Define SettingCard style for consistent content sections -->
        <Style TargetType="Border" x:Key="SettingCard">
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="4"/>
            <Setter Property="Padding" Value="12"/>
            <Setter Property="Margin" Value="0,0,0,15"/>
        </Style>
    </ResourceDictionary>
</Window.Resources>

<Border Background="{StaticResource WindowBackgroundBrush}" CornerRadius="12" BorderThickness="1" BorderBrush="{StaticResource BorderBrush}">
    <Border.Effect>
        <DropShadowEffect BlurRadius="20-30" ShadowDepth="10" Opacity="0.5" Color="Black"/>
    </Border.Effect>
    
    <!-- Window Content Goes Here (e.g. a Grid with Header, Content, Footer rows) -->
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Header -->
            <RowDefinition Height="*"/>    <!-- Content -->
            <RowDefinition Height="Auto"/> <!-- Footer -->
        </Grid.RowDefinitions>
        <!-- ... -->
    </Grid>
</Border>
```

### 3. Layout Patterns (Based on Settings Window)

#### **Header Pattern**
Use DockPanel for consistent header layout:
```xml
<!-- Header -->
<DockPanel Grid.Row="0" Margin="0,0,0,20">
    <TextBlock Text="Window Title" Style="{StaticResource DialogTitleText}"/>
    <Button x:Name="BtnClose" Style="{StaticResource WindowCloseButton}" Click="BtnCancel_Click"/>
</DockPanel>
```

#### **TabControl Pattern**
For complex windows with multiple sections, use TabControl:
```xml
<TabControl Grid.Row="1" Background="Transparent" BorderThickness="0,1,0,0" BorderBrush="#444">
    <TabControl.Resources>
        <Style TargetType="TabItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#AAA"/>
            <Setter Property="Padding" Value="16,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </TabControl.Resources>
    
    <TabItem Header="Tab Name">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
            <Grid Margin="10,10,20,20">
                <!-- Content here -->
            </Grid>
        </ScrollViewer>
    </TabItem>
</TabControl>
```

#### **Footer Pattern**
Use DockPanel for footer with action buttons:
```xml
<!-- Footer -->
<DockPanel Grid.Row="2" Margin="0,20,0,0" LastChildFill="False">
    <TextBlock Text="Label:" VerticalAlignment="Center" Margin="20,0,15,0" FontSize="13" Foreground="#AAA"/>
    <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
        <!-- Left side controls -->
    </StackPanel>
    <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
        <!-- Action buttons -->
        <Button Content="Cancel" Width="100" Height="36" 
                Background="{StaticResource SurfaceBrush}" BorderBrush="{StaticResource BorderBrush}" 
                Click="BtnCancel_Click" Margin="0,0,10,0"/>
        <Button x:Name="BtnOK" Content="OK" Width="120" Height="36" 
                Background="{StaticResource AccentBrush}" Foreground="White" FontWeight="SemiBold" 
                Click="BtnOK_Click"/>
    </StackPanel>
</DockPanel>
```

### 4. Content Organization

#### **SettingCard Pattern**
Use SettingCard style for content sections:
```xml
<Border Style="{StaticResource SettingCard}">
    <StackPanel>
        <TextBlock Text="Section Title" FontSize="16" FontWeight="SemiBold" Foreground="White" Margin="0,0,0,15"/>
        <!-- Section content -->
    </StackPanel>
</Border>
```

#### **Form Layout Pattern**
For forms with labels and inputs:
```xml
<Grid>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="100"/>
        <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="10"/>
        <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" Grid.Column="0" Text="Label" VerticalAlignment="Center" Foreground="#AAA"/>
    <TextBox x:Name="ControlName" Grid.Row="0" Grid.Column="1"/>
</Grid>
```

### 5. Drag-to-Move Behavior
Because `WindowStyle="None"` removes the native Windows title bar, you MUST implement dragging in the code-behind:
```csharp
private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
{
    if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
    {
        this.DragMove();
    }
}
```

### 6. Window Persistence vs. Auto-Closing

#### **Persistent Windows (Standard Utility Windows)**
Most utility windows (like Tag Inspector, Style Check, or Magic Resizer) should be **persistent**. They should stay open even when the user clicks back into PowerPoint to make adjustments.
- **Mandate:** DO NOT override `OnDeactivated` to call `this.Close()`.
- **Mandate:** Ensure `Topmost="True"` is set in XAML or `this.Topmost = true` in the constructor.
- **Mandate:** Set the PowerPoint window as the `Owner` using `WindowInteropHelper` in the constructor or feature trigger.

#### **Transient Windows (Context Menu Behavior)**
Only if the window acts as a quick context menu or selection picker (where interaction with the slide *while* the menu is open is not intended), use the `Deactivated` event to auto-close:
```xml
Deactivated="Window_Deactivated"
```
```csharp
private bool _isDialogOpened = false; // Flag to prevent close when child dialogs (like ColorDialog) open

private void Window_Deactivated(object sender, EventArgs e)
{
    if (!_isDialogOpened)
    {
        try { this.Close(); } catch { }
    }
}
```

### 7. UI Elements & Theme Guidelines

#### **Color Usage Rules**
- **NEVER hardcode colors** - Always use the centralized resource system
- **Headers/Main Text**: Use `{StaticResource TextPrimaryBrush}` for section headers and main titles
- **Secondary Text**: Use `{StaticResource TextSecondaryBrush}` or `{StaticResource TextMutedBrush}` for labels and secondary text
- **Theme Resources**: Use `{StaticResource ...}` for backgrounds, borders, accents
- **Backgrounds**: `{StaticResource WindowBackgroundBrush}` for main, `{StaticResource SurfaceBrush}` for cards
- **Accents**: `{StaticResource AccentBrush}` for highlights and primary actions
- **Borders**: `{StaticResource BorderBrush}` for standard borders, `"#444"` for tab separators

#### **Button Styling**
- **Corner Radius**: Use `CornerRadius="4"` (not pill-shaped)
- **Primary Actions**: `{StaticResource AccentBrush}` with `Foreground="White"`
- **Secondary Actions**: `{StaticResource SurfaceBrush}` with `{StaticResource BorderBrush}`
- **Always add**: `Cursor="Hand"`

#### **Control Styling**
- **TextBox**: Use default styling from Controls.xaml, add `Padding="5,2"` for consistency
- **ComboBox**: Standard elements - Controls.xaml handles dark theme automatically
- **CheckBox/RadioButton**: Default styling works well
- **ScrollViewer**: Use `VerticalScrollBarVisibility="Auto"` for scrollable content

#### **TreeView Styling (For Folder Trees)**
When displaying hierarchical folder structures, use proper contrast colors:
```xml
<TreeView x:Name="FolderTree" 
          BorderThickness="0" 
          Background="Transparent"
          Padding="0,4"
          SelectedItemChanged="FolderTree_SelectedItemChanged">
    <TreeView.Resources>
        <HierarchicalDataTemplate DataType="{x:Type features:FolderTreeNode}" ItemsSource="{Binding Children}">
            <StackPanel Orientation="Horizontal" Margin="2,1" ToolTip="{Binding FullPath}">
                <!-- Icons: Use #AAAAAA (light grey) -->
                <TextBlock Text="{Binding Icon}" FontSize="12" Margin="0,0,5,0" 
                           VerticalAlignment="Center" Foreground="#AAAAAA"/>
                <!-- Text: Use #E0E0E0 (almost white) for readability on dark background -->
                <TextBlock Text="{Binding DisplayName}" FontSize="12" 
                           VerticalAlignment="Center" Foreground="#E0E0E0"/>
            </StackPanel>
        </HierarchicalDataTemplate>
    </TreeView.Resources>
    <TreeView.ItemContainerStyle>
        <Style TargetType="TreeViewItem">
            <Setter Property="IsExpanded" Value="True"/>
            <Setter Property="Padding" Value="3,2"/>
            <Setter Property="Margin" Value="0"/>
            <Setter Property="Foreground" Value="#CCCCCC"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#0078D4"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </TreeView.ItemContainerStyle>
</TreeView>
```
**Important**: The project theme provides `StaticResource` brushes like `{StaticResource WindowBackgroundBrush}` and `{StaticResource SurfaceBrush}` across all windows. For TreeView text on dark backgrounds, always use hardcoded light colors:
- **Folder icons**: `#AAAAAA` (muted light grey)
- **Text labels**: `#E0E0E0` (readable light grey)
- **Selected state**: `#0078D4` background with `White` foreground

#### **Typography Guidelines**
- **Window Titles**: `Style="{StaticResource DialogTitleText}"`
- **Section Headers**: `Style="{StaticResource DialogSubtitleText}"`
- **Sub-Headers**: `FontSize="12" FontWeight="Bold" Foreground="{StaticResource AccentBrush}"`
- **Labels**: `FontSize="11" Foreground="{StaticResource TextSecondaryBrush}"`
- **Content Text**: Default styling from Controls.xaml

#### **Spacing & Layout**
- **Main Margins**: `Margin="20"` for main grid
- **Content Margins**: `Margin="10,10,20,20"` for tab content
- **Section Spacing**: `Margin="0,0,0,15"` between SettingCards
- **Control Spacing**: `Margin="0,0,0,10"` between form controls
- **Header Margin**: `Margin="0,0,0,20"` below header

### 8. Window Sizing Guidelines
- **Simple Windows**: `Height="450-550" Width="600-700"`
- **Complex Windows**: `Height="650-750" Width="800-900"`
- **Always add**: ScrollViewer for content that might overflow

## Execution Steps for the AI Agent
When the user asks to "create a floating window for X":
1. Create `UI/XWindow.xaml` and `UI/XWindow.xaml.cs`
2. Apply the above `Window` properties and use proper font family
3. Include the `ResourceDictionary.MergedDictionaries` with converters and styles
4. Apply the curved `Border` style with appropriate shadow settings
5. Use proper layout patterns (DockPanel for header/footer, TabControl for complex content)
6. Apply color guidelines (White for headers, #AAA for labels, StaticResource for theme colors)
7. Use SettingCard style for content organization
8. Implement proper button styling with CornerRadius="4"
9. Add ScrollViewer for scrollable content areas
10. Implement `Window_MouseDown` for dragging
11. If it needs to act on a PowerPoint shape, pass `PowerPoint.Shape` via the constructor
12. Test that all text is readable and follows the dark theme patterns
13. **IMPORTANT:** If the window is part of a new feature, ensure the feature is added to the `FeatureLibrary.AllFeatures` static list in **alphabetical order by Name** so it appears correctly in the Settings menu.
14. **IMPORTANT:** If the floating window is a tool/palette that should only exist once at any given time, register it as `IsToggle = true` in the `FeatureLibrary` and use the `ToggleFloatingWindow` helper in `MainSidebar.xaml.cs` to launch it.
order` style with appropriate shadow settings
5. Use proper layout patterns (DockPanel for header/footer, TabControl for complex content)
6. Apply color guidelines (White for headers, #AAA for labels, StaticResource for theme colors)
7. Use SettingCard style for content organization
8. Implement proper button styling with CornerRadius="4"
9. Add ScrollViewer for scrollable content areas
10. Implement `Window_MouseDown` for dragging
11. If it needs to act on a PowerPoint shape, pass `PowerPoint.Shape` via the constructor
12. Test that all text is readable and follows the dark theme patterns
13. **IMPORTANT:** If the window is part of a new feature, ensure the feature is added to the `FeatureLibrary.AllFeatures` static list in **alphabetical order by Name** so it appears correctly in the Settings menu.
