package pe.edu.upc.dsd.neteros;

import java.util.ArrayList;
import java.util.List;

public class ClientesServImpl implements ClienteService {
private List<Cliente> clientes = new ArrayList<Cliente>();
	
		@Override
		public void registrar(Cliente cliente) {
			clientes.add(cliente);			
		}
}
