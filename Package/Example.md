# Pure Media Player 自定义解码包使用说明
## 0.自由定制解码包

[范例文件(Package.ini)](Default/Package.ini)
(*注意：该文件不是标准的INI文件*)

>本软件允许用户完全自定义解码包，支持单向的DirectShow IPin接口连接。
>用户可以通过自定义解码插件与Package.ini文件来实现解码器自定义。
>除了使用DirectShow自带的SourceFilter滤镜之外，其他任何滤镜都可以由用户自定义。
>用户无需过多的处理具体链接的代码，只需要为程序提供链接清单即可。  
> 本软件默认自带的是LAVFilters里的Splitter、Video、Audio三款开放源代码的解码器与节流器、vsfilter的DirectVobSub字幕插件和matashi的madVR渲染器。

## 1.头属性

> `Name=Default`  
>定义解码包在程序中显示的名字，可以为任意字符串。这里为Default。  
> `Creator=Deliay`  
>定义解码包在程序中显示的作者名字。

>`MediaSource=Splitter`  
>定义程序SourceFilter将会把Input接口传给哪个Filter。这里为Splitter。  
>`MediaSourcePin=Input`  
>定义程序SourceFilter将会把Input接口传给Filter的哪个Pin。这里为Input。

## 2.滤镜列表声明
>`[Section]`  
>Section段，表明这是一个滤镜声明的字段，指明了一共哪几个滤镜将会被加载到程序中  
>`=@`  
>用于标识这是一个列表（需要写上）

>`=CustomFilterName1`  
`=CustomFilterName2`  
`=CustomFilterName3`  
>`=………………`  
>这里写自定义滤镜名称的列表，这里只是作为一个 **标识符** 存在，不需要实际的滤镜名称作为关联(可以任意起名字)

## 3.滤镜声明
>`[CustomFilterName1]`  
>指明这是哪个自定义滤镜的声明  
>`@=@`  
>指明这是一个HashMap，需要写上

>`#File=Filter.ax`  
>指明滤镜文件，可以是绝对定位路径，也可以是相对定位路径。  
>`#CLSID=00000000-0000-0000-0000-000000000000`  
>指明需要调用的滤镜的CLSID，会根据这个从所指定的滤镜文件中取得滤镜的实例。必须有效。

>`~Video=CustomFilterName2`  
>将CustomFilterName1的Video链接到CustomFilterName2  
>`VideoPin=Input`  
>链接到CustomFilterName2的Input这个Pin上

>`#Render=Output`  
>使用系统默认设置对接下来的Pin进行连接。