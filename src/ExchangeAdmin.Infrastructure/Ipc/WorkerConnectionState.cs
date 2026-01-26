namespace ExchangeAdmin.Infrastructure.Ipc;

             
                                      
              
public enum WorkerConnectionState
{
                 
                           
                  
    NotStarted,

                 
                                
                  
    Starting,

                 
                                            
                  
    WaitingForHandshake,

                 
                                 
                  
    Connected,

                 
                          
                  
    Restarting,

                 
                                   
                  
    Stopped,

                 
                        
                  
    Crashed,

                 
                                                
                  
    Unresponsive
}
