# UI Layout Rules

## WinForms 设置区防重叠规则

1. 设置项禁止用“标签固定 X + 输入框固定 X”的方式硬排。
2. 每个设置项必须用一个容器包住标签和控件，标签 `AutoSize = true`，控件紧跟标签后面。
3. 设置项之间必须用容器 `Margin` 留间距，不能靠猜测中文标签宽度。
4. 同一行设置项优先使用 `FlowLayoutPanel` 或 `TableLayoutPanel` 自动排版。
5. 新增中文较长标签后，必须检查最小窗口宽度下是否与输入框、复选框、按钮重叠。
6. 按钮可以靠右动态定位，但设置项区域必须给按钮预留独立空间，不能和按钮共用硬编码坐标。

示例：

```csharp
var row = new FlowLayoutPanel
{
    FlowDirection = FlowDirection.LeftToRight,
    WrapContents = false
};

row.Controls.Add(CreateLabeledOption("图形纵行间距(mm):", txtNodeSpacing));
```
