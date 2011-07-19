' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de interfaz "IService1" en el código y en el archivo de configuración a la vez.
Imports System.ServiceModel
Imports System.Runtime.Serialization
Imports System.Data

<ServiceContract()>
Public Interface IServiceGroupon
    <OperationContract()>
    Function RegistraCliente(ByVal oCliente As String) As Integer

End Interface

' Utilice un contrato de datos, como se ilustra en el ejemplo siguiente, para agregar tipos compuestos a las operaciones de servicio.

<DataContract()>
Public Class CompositeType

    <DataMember()>
    Public Property BoolValue() As Boolean

    <DataMember()>
    Public Property StringValue() As String

End Class
