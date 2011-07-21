' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de interfaz "IService1" en el código y en el archivo de configuración a la vez.
Imports System.ServiceModel
Imports System.Runtime.Serialization
Imports System.Data

<ServiceContract()>
Public Interface IServiceGroupon
    <OperationContract()>
    Function RegistraCliente(ByVal oCliente As String) As Integer

    <OperationContract()>
    Function ValidarUsuario(ByVal pcMail As String) As String

    <OperationContract()>
    Function RegistrarCupon(ByVal oCupon As String, ByVal oFoto() As Byte) As Integer

    <OperationContract()>
    Function DevuelveOfertas() As String

    <OperationContract()>
    Function RegistrarPagoCupon(ByVal oPagos As String, ByVal pcMail As String) As Integer



    ' Utilice un contrato de datos, como se ilustra en el ejemplo siguiente, para agregar tipos compuestos a las operaciones de servicio.
    <DataContract()>
    Class ClsCupon
        <DataMember()> Public cTitulo As String
        <DataMember()> Public cDescripcion As String
        <DataMember()> Public nCosto As Double
        <DataMember()> Public nDesc As Double
        <DataMember()> Public oFotoOferta() As Byte
    End Class
End Interface

