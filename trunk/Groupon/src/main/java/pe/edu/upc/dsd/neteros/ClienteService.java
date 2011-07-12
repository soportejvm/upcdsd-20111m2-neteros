package pe.edu.upc.dsd.neteros;    
import javax.jws.WebParam;
import javax.jws.WebService;    

@WebService  
public interface ClienteService {            
	public void registrar(@WebParam(name = "cliente") Cliente cliente); 
	
}