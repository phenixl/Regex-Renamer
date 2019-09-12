Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports phx



'''<summary>
'''这是 MainModuleTest 的测试类，旨在
'''包含所有 MainModuleTest 单元测试
'''</summary>
<TestClass()> _
Public Class MainModuleTest


    Private testContextInstance As TestContext

    '''<summary>
    '''获取或设置测试上下文，上下文提供
    '''有关当前测试运行及其功能的信息。
    '''</summary>
    Public Property TestContext() As TestContext
        Get
            Return testContextInstance
        End Get
        Set(ByVal value As TestContext)
            testContextInstance = Value
        End Set
    End Property

#Region "附加测试特性"
    '
    '编写测试时，还可使用以下特性:
    '
    '使用 ClassInitialize 在运行类中的第一个测试前先运行代码
    '<ClassInitialize()>  _
    'Public Shared Sub MyClassInitialize(ByVal testContext As TestContext)
    'End Sub
    '
    '使用 ClassCleanup 在运行完类中的所有测试后再运行代码
    '<ClassCleanup()>  _
    'Public Shared Sub MyClassCleanup()
    'End Sub
    '
    '使用 TestInitialize 在运行每个测试前先运行代码
    '<TestInitialize()>  _
    'Public Sub MyTestInitialize()
    'End Sub
    '
    '使用 TestCleanup 在运行完每个测试后运行代码
    '<TestCleanup()>  _
    'Public Sub MyTestCleanup()
    'End Sub
    '
#End Region


    '''<summary>
    '''PrintHelp 的测试
    '''</summary>
    <TestMethod(), _
     DeploymentItem("rren.exe")> _
    Public Sub PrintHelpTest()
        MainModule_Accessor.PrintHelp()
        Assert.Inconclusive("无法验证不返回值的方法。")
    End Sub

    '''<summary>
    '''Print 的测试
    '''</summary>
    <TestMethod()> _
    Public Sub PrintTest()
        Dim msg As String = String.Empty ' TODO: 初始化为适当的值
        MainModule.Print(msg)
        Assert.Inconclusive("无法验证不返回值的方法。")
    End Sub
End Class
