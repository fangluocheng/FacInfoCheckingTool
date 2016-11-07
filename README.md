## 电视工厂信息校验工具

此工具通过串口发送命令给电视，获取电视返回的数据。并且与设定好的数据进行比较。如果与设定好的数据一样，那么就通过测试。否则显示 NG 并提示错误的数据。

### 运行须知

生成的 \*.exe 程序需要与 **facInfoData.mdb** 数据库文件和 **MSCOMM32.OCX** 文件放在同一个目录下才可以运行成功。

#### 注册 MSCOMM32.OCX 和 MSWINSCK.OCX

由于此工具用到了串口通讯和网口通讯，所以在使用之前，需要确保电视注册了 MSCOMM32.OCX 和 MSWINSCK.OCX 两个控件。

对于 Windows 7 64 位系统，按照如下步骤进行注册：

1. 如果系统中没有 [MSCOMM32.OCX](https://github.com/heray1990/FacInfoCheckingTool/blob/master/MSCOMM32.OCX) 和 [MSWINSCK.OCX](https://github.com/heray1990/FacInfoCheckingTool/blob/master/MSWINSCK.OCX) 这两个控件，需要将这两个控件复制到 `C:\Windows\SysWOW64` 文件夹中。
2. 进入 `C:\Windows\SysWOW64` 文件夹，找到 `cmd.exe` 这个文件。右击，选择“以管理员身份运行”。
3. 按照下面的截图输入两条命令注册上述两个控件。

![register_ocx_in_windows7_64](https://github.com/heray1990/FacInfoCheckingTool/raw/master/Images/register_ocx_in_windows7_64.png)

> Note: 上述的 `C:\Windows\SysWOW64` 文件夹是针对 Windows 64 位系统的，对于 32 位系统，需要改成 `C:\Windows\System32`。