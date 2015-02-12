Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports MailStewardControl
Imports System.Drawing



'''<summary>
'''This is a test class for ucPadTest and is intended
'''to contain all ucPadTest Unit Tests
'''</summary>
<TestClass()> _
Public Class ucPadTest


    Private testContextInstance As TestContext

    '''<summary>
    '''Gets or sets the test context which provides
    '''information about and functionality for the current test run.
    '''</summary>
    Public Property TestContext() As TestContext
        Get
            Return testContextInstance
        End Get
        Set(ByVal value As TestContext)
            testContextInstance = Value
        End Set
    End Property

#Region "Additional test attributes"
    '
    'You can use the following additional attributes as you write your tests:
    '
    'Use ClassInitialize to run code before running the first test in the class
    '<ClassInitialize()>  _
    'Public Shared Sub MyClassInitialize(ByVal testContext As TestContext)
    'End Sub
    '
    'Use ClassCleanup to run code after all tests in a class have run
    '<ClassCleanup()>  _
    'Public Shared Sub MyClassCleanup()
    'End Sub
    '
    'Use TestInitialize to run code before running each test
    '<TestInitialize()>  _
    'Public Sub MyTestInitialize()
    'End Sub
    '
    'Use TestCleanup to run code after each test has run
    '<TestCleanup()>  _
    'Public Sub MyTestCleanup()
    'End Sub
    '
#End Region


    '''<summary>
    '''A test for ProcessMail
    '''</summary>
    <TestMethod()> _
    Public Sub ProcessMailTest()
        Dim frm As New TestPadForm
        Dim target As ucPad = New ucPad ' TODO: Initialize to an appropriate value
        frm.Location = New Point(0, 0)
        frm.Controls.Add(target)
        frm.Show()
        Dim sSubject As String = String.Empty ' TODO: Initialize to an appropriate value
        target.ProcessMail(sSubject)
        MsgBox("OK")
        Assert.Inconclusive("A method that does not return a value cannot be verified.")
    End Sub
End Class
