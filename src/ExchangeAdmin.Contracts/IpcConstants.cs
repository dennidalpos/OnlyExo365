namespace ExchangeAdmin.Contracts;

             
                                                                
                                                                                  
              
public static class IpcConstants
{
    #region Pipe Names

                 
                                                               
                  
    public const string PipeName = "ExchangeAdmin_IPC_Main";

                 
                                                                        
                  
    public const string EventPipeName = "ExchangeAdmin_IPC_Events";

    #endregion

    #region Timeouts

                 
                                                                                   
                  
    public const int ConnectionTimeoutMs = 10000;

                 
                                                                         
                  
    public const int HandshakeTimeoutMs = 5000;

                 
                                                                                                           
                  
    public const int RequestTimeoutMs = 300000;            

                 
                                           
                  
    public const int HeartbeatIntervalMs = 5000;

                 
                                                                                                     
                  
    public const int HeartbeatTimeoutMs = 15000;

                 
                                                                                     
                                                 
                  
    public const int HeartbeatGracePeriodMs = 5000;

                 
                                                                         
                  
    public const int HeartbeatMissedThreshold = 3;

    #endregion

    #region Buffer & Limits

                 
                                        
                  
    public const int PipeBufferSize = 65536;

                 
                                                                         
                                                                  
                  
    public const int MaxMessageSizeBytes = 10 * 1024 * 1024;

                 
                                                               
                                                           
                  
    public const int MaxEventsPerRequest = 10000;

                 
                                                                                
                  
    public const int MaxReadBufferSize = 256 * 1024;

    #endregion

    #region Protocol

                 
                                                         
                                                             
                  
    public const char MessageDelimiter = '\n';

                 
                                                  
                  
    public const int ProtocolVersionMajor = 1;

    #endregion

    #region Validation

                 
                                                       
                  
                                                            
                                                                
    public static bool IsValidMessageSize(long sizeBytes)
        => sizeBytes > 0 && sizeBytes <= MaxMessageSizeBytes;

                 
                                                          
                  
                                                          
                                                   
    public static bool IsEventCountWithinLimit(int eventCount)
        => eventCount >= 0 && eventCount < MaxEventsPerRequest;

    #endregion
}
