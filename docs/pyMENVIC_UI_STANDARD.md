# pyMENVIC UI Standard

Referencia visual base para unificar interfaces XAML/WPF de herramientas pyMENVIC.

Base tomada de la pantalla estable:

`pyMenvic.tab/Filters.panel/FilterManagerPro.pushbutton/filter_manager_pro.xaml`

## Principios

- Aplicar cambios quirurgicos por herramienta.
- No redisenar una pantalla completa si ya funciona.
- Preservar nombres de controles, bindings y eventos existentes.
- No tocar logica Python cuando el objetivo sea solo UI.
- Mantener compatibilidad WPF + IronPython + pyRevit.

## Paleta base aprobada

### Ventana y estructura

| Uso | Color |
|---|---|
| Window background | `#1E252B` |
| Main grid background | `#1E252B` |
| Header background | `#26303A` |
| Content background | `#F4F6F8` |
| Header cards background | `#1F2B35` |
| Header cards border | `#344B5B` |
| Footer text | `#A9BAC8` |

### Texto

| Uso | Color |
|---|---|
| Primary dark text | `#1E252B` |
| Primary light text | `#FFFFFF` |
| Header card label | `#8FA7BA` |
| Status blue text | `#0A70B8` |

### Botones

#### ToolButtonStyle

| Propiedad | Valor |
|---|---|
| Foreground | `White` |
| Background | `#2F4B5E` |
| BorderBrush | `#4C6B82` |
| Height | `30` |
| Padding | `12,4` |
| Margin | `0,0,6,0` |

#### ApplyButtonStyle

| Propiedad | Valor |
|---|---|
| Background | `#2D9D55` |
| BorderBrush | `#49B571` |

#### DangerButtonStyle

| Propiedad | Valor |
|---|---|
| Background | `#7A3430` |
| BorderBrush | `#9A4A45` |

## Inputs

### SectionLabelStyle

| Propiedad | Valor |
|---|---|
| Foreground | `#1E252B` |
| FontWeight | `SemiBold` |
| Margin | `0,0,6,0` |
| VerticalAlignment | `Center` |

### ValueBoxStyle

| Propiedad | Valor |
|---|---|
| Background | `#F7F9FB` |
| Foreground | `#1E252B` |
| BorderBrush | `#9FB3C1` |
| Height | `30` |
| Padding | `6,4` |
| VerticalContentAlignment | `Center` |

### ComboBox base

| Propiedad | Valor |
|---|---|
| Background | `#F7F9FB` |
| Foreground | `#1E252B` |
| Height | `30` |
| VerticalContentAlignment | `Center` |

### ComboBoxItem base

| Propiedad | Valor |
|---|---|
| Background | `#FFFFFF` |
| Foreground | `#1E252B` |

## DataGrid

| Propiedad | Valor |
|---|---|
| CanUserAddRows | `False` |
| Background | `#F7F9FB` |
| Foreground | `#1E252B` |
| RowBackground | `#FFFFFF` |
| AlternatingRowBackground | `#F0F3F6` |
| GridLinesVisibility | `All` |
| HorizontalGridLinesBrush | `#C8D1D8` |
| VerticalGridLinesBrush | `#C8D1D8` |
| RowHeight | `22` |
| ColumnHeaderHeight | `26` |
| SelectionMode | `Single` |
| SelectionUnit | `FullRow` |

### DataGridRow selected

| Propiedad | Valor |
|---|---|
| Background | `#0A70B8` |
| Foreground | `White` |
| FontWeight | `SemiBold` |

### DataGridCell selected

| Propiedad | Valor |
|---|---|
| Background | `#0A70B8` |
| Foreground | `White` |

### DataGridColumnHeader

| Propiedad | Valor |
|---|---|
| Background | `#E9EEF3` |
| Foreground | `#1E252B` |
| FontWeight | `SemiBold` |
| Padding | `6,4` |

## Tabs

### Inactive TabItem

| Propiedad | Valor |
|---|---|
| Foreground | `#C3D1DD` |
| Background | `#24323E` |
| Padding | `14,8` |

### Hover TabItem

| Propiedad | Valor |
|---|---|
| Foreground | `#1E252B` |
| Background | `#DCE7EF` |
| FontWeight | `SemiBold` |

### Selected TabItem

| Propiedad | Valor |
|---|---|
| Foreground | `#1E252B` |
| Background | `#F4F6F8` |
| FontWeight | `SemiBold` |

## Header

Base recomendada:

- Header background: `#26303A`
- Padding: `10`
- Logo: `28 x 28`
- Title text: white, size `18`, bold
- Cards background: `#1F2B35`
- Cards border: `#344B5B`
- Cards corner radius: `4`
- Cards label: `#8FA7BA`, size `11`
- Cards value: white, size `16`, bold

## Footer

Base recomendada:

- Background: `#1E252B`
- BorderThickness: `0`
- Padding: `3,4,3,1`
- Margin: `0,2,0,0`
- Version text foreground: `#A9BAC8`
- FontSize: `12`
- HorizontalAlignment: `Left`

Formato recomendado de version:

`pyMENVIC <Tool Name> | MVP 0.x.x`

## Bloque XAML base de recursos

```xml
<Window.Resources>
  <Style x:Key="ToolButtonStyle" TargetType="Button">
    <Setter Property="Foreground" Value="White"/>
    <Setter Property="Background" Value="#2F4B5E"/>
    <Setter Property="BorderBrush" Value="#4C6B82"/>
    <Setter Property="Height" Value="30"/>
    <Setter Property="Padding" Value="12,4"/>
    <Setter Property="Margin" Value="0,0,6,0"/>
  </Style>
  <Style x:Key="ApplyButtonStyle" TargetType="Button" BasedOn="{StaticResource ToolButtonStyle}">
    <Setter Property="Background" Value="#2D9D55"/>
    <Setter Property="BorderBrush" Value="#49B571"/>
  </Style>
  <Style x:Key="DangerButtonStyle" TargetType="Button" BasedOn="{StaticResource ToolButtonStyle}">
    <Setter Property="Background" Value="#7A3430"/>
    <Setter Property="BorderBrush" Value="#9A4A45"/>
  </Style>
  <Style x:Key="SectionLabelStyle" TargetType="TextBlock">
    <Setter Property="Foreground" Value="#1E252B"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Margin" Value="0,0,6,0"/>
    <Setter Property="VerticalAlignment" Value="Center"/>
  </Style>
  <Style x:Key="ValueBoxStyle" TargetType="TextBox">
    <Setter Property="Background" Value="#F7F9FB"/>
    <Setter Property="Foreground" Value="#1E252B"/>
    <Setter Property="BorderBrush" Value="#9FB3C1"/>
    <Setter Property="Height" Value="30"/>
    <Setter Property="Padding" Value="6,4"/>
    <Setter Property="VerticalContentAlignment" Value="Center"/>
  </Style>
  <Style TargetType="ComboBox">
    <Setter Property="Background" Value="#F7F9FB"/>
    <Setter Property="Foreground" Value="#1E252B"/>
    <Setter Property="Height" Value="30"/>
    <Setter Property="VerticalContentAlignment" Value="Center"/>
  </Style>
  <Style TargetType="ComboBoxItem">
    <Setter Property="Background" Value="#FFFFFF"/>
    <Setter Property="Foreground" Value="#1E252B"/>
  </Style>
  <Style TargetType="DataGrid">
    <Setter Property="CanUserAddRows" Value="False"/>
    <Setter Property="Background" Value="#F7F9FB"/>
    <Setter Property="Foreground" Value="#1E252B"/>
    <Setter Property="RowBackground" Value="#FFFFFF"/>
    <Setter Property="AlternatingRowBackground" Value="#F0F3F6"/>
    <Setter Property="GridLinesVisibility" Value="All"/>
    <Setter Property="HorizontalGridLinesBrush" Value="#C8D1D8"/>
    <Setter Property="VerticalGridLinesBrush" Value="#C8D1D8"/>
    <Setter Property="RowHeight" Value="22"/>
    <Setter Property="ColumnHeaderHeight" Value="26"/>
    <Setter Property="SelectionMode" Value="Single"/>
    <Setter Property="SelectionUnit" Value="FullRow"/>
  </Style>
  <Style TargetType="DataGridRow">
    <Setter Property="Foreground" Value="#1E252B"/>
    <Style.Triggers>
      <Trigger Property="IsSelected" Value="True">
        <Setter Property="Background" Value="#0A70B8"/>
        <Setter Property="Foreground" Value="White"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style TargetType="DataGridCell">
    <Setter Property="Foreground" Value="#1E252B"/>
    <Setter Property="BorderBrush" Value="#C8D1D8"/>
    <Style.Triggers>
      <Trigger Property="IsSelected" Value="True">
        <Setter Property="Background" Value="#0A70B8"/>
        <Setter Property="Foreground" Value="White"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style TargetType="DataGridColumnHeader">
    <Setter Property="Background" Value="#E9EEF3"/>
    <Setter Property="Foreground" Value="#1E252B"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Padding" Value="6,4"/>
  </Style>
  <Style TargetType="TabItem">
    <Setter Property="Foreground" Value="#C3D1DD"/>
    <Setter Property="Background" Value="#24323E"/>
    <Setter Property="Padding" Value="14,8"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="Foreground" Value="#1E252B"/>
        <Setter Property="Background" Value="#DCE7EF"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
      </Trigger>
      <Trigger Property="IsSelected" Value="True">
        <Setter Property="Foreground" Value="#1E252B"/>
        <Setter Property="Background" Value="#F4F6F8"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
      </Trigger>
    </Style.Triggers>
  </Style>
</Window.Resources>
```

## Checklist antes de aplicar a una herramienta

- Revisar XAML actual.
- Buscar referencias de controles en `script.py`.
- No renombrar controles sin revisar uso en Python.
- Aplicar primero recursos globales: botones, DataGrid, tabs, TextBox, ComboBox.
- Ajustar solo colores/margenes puntuales si la pantalla lo necesita.
- Probar carga de la ventana.
- Probar botones principales.
- Commit pequeno por herramienta.
