using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

             
                                                                          
                                                                                          
                                               
              
             
          
                                                                                   
                                                                  
           
          
                           
                        
                                                                                      
                                                                                           
           
           
              
public class CapabilityMapDto
{
                 
                                                                  
                                                          
                  
    [JsonPropertyName("detectedAt")]
    public DateTime DetectedAt { get; set; } = DateTime.UtcNow;

                 
                                                           
                                                                          
                  
                 
                                                                                   
                  
    [JsonPropertyName("cmdlets")]
    public Dictionary<string, CmdletCapabilityDto> Cmdlets { get; set; } = new();

                 
                                                                             
                                                                 
                  
    [JsonPropertyName("features")]
    public FeatureCapabilitiesDto Features { get; set; } = new();
}

             
                                                                           
              
public class CmdletCapabilityDto
{
                 
                                                                                
                  
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

                 
                                                                                  
                  
    [JsonPropertyName("isAvailable")]
    public bool IsAvailable { get; set; }

                 
                                                  
                                        
                  
    [JsonPropertyName("parameters")]
    public List<string> Parameters { get; set; } = new();

                 
                                                                          
                                                                                    
                  
    [JsonPropertyName("unavailableReason")]
    public string? UnavailableReason { get; set; }
}

             
                                                               
                                                                    
              
             
          
                              
                       
                                                                                         
                                                                                                           
                                                                                                                        
                                                                                                                                                
                                                                                                                                                                
                                                                                                                                           
                                                                                                               
           
           
              
public class FeatureCapabilitiesDto
{
    #region Mailbox Read Operations

                 
                                                             
                                                       
                  
    [JsonPropertyName("canGetMailbox")]
    public bool CanGetMailbox { get; set; }

                 
                                                                              
                                                           
                  
    [JsonPropertyName("canGetMailboxStatistics")]
    public bool CanGetMailboxStatistics { get; set; }

                 
                                                            
                                                                    
                  
    [JsonPropertyName("canGetInboxRule")]
    public bool CanGetInboxRule { get; set; }

                 
                                                                       
                                                           
                  
    [JsonPropertyName("canGetMailboxAutoReplyConfiguration")]
    public bool CanGetMailboxAutoReplyConfiguration { get; set; }

    #endregion

    #region Mailbox Write Operations

                 
                                                              
                                                   
                  
    [JsonPropertyName("canSetMailbox")]
    public bool CanSetMailbox { get; set; }

                 
                                                                                  
                                                             
                  
    [JsonPropertyName("canSetArchive")]
    public bool CanSetArchive { get; set; }

                 
                                                                  
                                                                    
                  
    [JsonPropertyName("canSetLitigationHold")]
    public bool CanSetLitigationHold { get; set; }

                 
                                                         
                                                    
                  
    [JsonPropertyName("canSetAudit")]
    public bool CanSetAudit { get; set; }

    #endregion

    #region Permission Operations

                 
                                                                  
                                                           
                  
    [JsonPropertyName("canGetMailboxPermission")]
    public bool CanGetMailboxPermission { get; set; }

                 
                                                     
                                                    
                  
    [JsonPropertyName("canAddMailboxPermission")]
    public bool CanAddMailboxPermission { get; set; }

                 
                                                        
                                                     
                  
    [JsonPropertyName("canRemoveMailboxPermission")]
    public bool CanRemoveMailboxPermission { get; set; }

                 
                                                                
                                                       
                  
    [JsonPropertyName("canGetRecipientPermission")]
    public bool CanGetRecipientPermission { get; set; }

                 
                                                       
                                                
                  
    [JsonPropertyName("canAddRecipientPermission")]
    public bool CanAddRecipientPermission { get; set; }

                 
                                                          
                                                 
                  
    [JsonPropertyName("canRemoveRecipientPermission")]
    public bool CanRemoveRecipientPermission { get; set; }

    #endregion

    #region Distribution List Operations

                 
                                                     
                                                                
                  
    [JsonPropertyName("canGetDistributionGroup")]
    public bool CanGetDistributionGroup { get; set; }

                 
                                                     
                                                                   
                  
    [JsonPropertyName("canSetDistributionGroup")]
    public bool CanSetDistributionGroup { get; set; }

    [JsonPropertyName("canSetDistributionGroupRequireSenderAuthentication")]
    public bool CanSetDistributionGroupRequireSenderAuthentication { get; set; }

    [JsonPropertyName("canSetDistributionGroupAcceptMessagesOnlyFrom")]
    public bool CanSetDistributionGroupAcceptMessagesOnlyFrom { get; set; }

    [JsonPropertyName("canSetDistributionGroupRejectMessagesFrom")]
    public bool CanSetDistributionGroupRejectMessagesFrom { get; set; }

                 
                                                           
                                                     
                  
    [JsonPropertyName("canGetDistributionGroupMember")]
    public bool CanGetDistributionGroupMember { get; set; }

                 
                                                           
                                                        
                  
    [JsonPropertyName("canAddDistributionGroupMember")]
    public bool CanAddDistributionGroupMember { get; set; }

                 
                                                              
                                                          
                  
    [JsonPropertyName("canRemoveDistributionGroupMember")]
    public bool CanRemoveDistributionGroupMember { get; set; }

    #endregion

    #region Dynamic Distribution Group Operations

                 
                                                            
                                            
                  
    [JsonPropertyName("canGetDynamicDistributionGroup")]
    public bool CanGetDynamicDistributionGroup { get; set; }

    [JsonPropertyName("canSetDynamicDistributionGroup")]
    public bool CanSetDynamicDistributionGroup { get; set; }

    [JsonPropertyName("canSetDynamicDistributionGroupRequireSenderAuthentication")]
    public bool CanSetDynamicDistributionGroupRequireSenderAuthentication { get; set; }

    [JsonPropertyName("canSetDynamicDistributionGroupAcceptMessagesOnlyFrom")]
    public bool CanSetDynamicDistributionGroupAcceptMessagesOnlyFrom { get; set; }

    [JsonPropertyName("canSetDynamicDistributionGroupRejectMessagesFrom")]
    public bool CanSetDynamicDistributionGroupRejectMessagesFrom { get; set; }

                 
                                                                  
                                                              
                  
    [JsonPropertyName("canGetDynamicDistributionGroupMember")]
    public bool CanGetDynamicDistributionGroupMember { get; set; }

    #endregion

    #region Microsoft 365 Group Operations

                 
                                                
                                         
                  
    [JsonPropertyName("canGetUnifiedGroup")]
    public bool CanGetUnifiedGroup { get; set; }

                 
                                                     
                                                                
                  
    [JsonPropertyName("canGetUnifiedGroupLinks")]
    public bool CanGetUnifiedGroupLinks { get; set; }

    #endregion
}

             
                                      
              
public class DetectCapabilitiesRequest
{
                 
                                                                                   
                  
    [JsonPropertyName("forceRefresh")]
    public bool ForceRefresh { get; set; }
}
