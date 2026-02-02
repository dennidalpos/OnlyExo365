namespace ExchangeAdmin.Contracts.Messages;

             
                          
              
public enum MessageType
{
                
    HandshakeRequest,
    HandshakeResponse,

                       
    Request,
    Response,

                       
    Event,

              
    CancelRequest,
    HeartbeatPing,
    HeartbeatPong
}
