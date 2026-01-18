using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

/// <summary>
/// Mappa delle capability rilevate dopo la connessione a Exchange Online.
/// Contiene informazioni su quali cmdlet sono disponibili e quali feature sono supportate
/// in base ai ruoli RBAC dell'utente connesso.
/// </summary>
/// <remarks>
/// <para>
/// La capability map viene rilevata automaticamente dopo ogni connessione riuscita
/// ed è usata dalla UI per abilitare/disabilitare funzionalità.
/// </para>
/// <para>
/// <b>Capability Keys:</b>
/// <list type="bullet">
/// <item><c>Cmdlets</c>: Dictionary keyed by cmdlet name (e.g., "Get-Mailbox")</item>
/// <item><c>Features</c>: Boolean flags per feature aggregate (e.g., CanGetMailbox)</item>
/// </list>
/// </para>
/// </remarks>
public class CapabilityMapDto
{
    /// <summary>
    /// Timestamp UTC di quando le capability sono state rilevate.
    /// Usato per determinare se è necessario un refresh.
    /// </summary>
    [JsonPropertyName("detectedAt")]
    public DateTime DetectedAt { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Mappa dei cmdlet disponibili, keyed by nome cmdlet.
    /// Esempio: "Get-Mailbox" -> { IsAvailable: true, Parameters: [...] }
    /// </summary>
    /// <remarks>
    /// Include solo i cmdlet rilevanti per l'applicazione, non tutti i cmdlet EXO.
    /// </remarks>
    [JsonPropertyName("cmdlets")]
    public Dictionary<string, CmdletCapabilityDto> Cmdlets { get; set; } = new();

    /// <summary>
    /// Flag booleani per feature aggregate, derivati dai cmdlet disponibili.
    /// Usati dalla UI per abilitare/disabilitare sezioni intere.
    /// </summary>
    [JsonPropertyName("features")]
    public FeatureCapabilitiesDto Features { get; set; } = new();
}

/// <summary>
/// Informazioni sulla disponibilità di un singolo cmdlet Exchange Online.
/// </summary>
public class CmdletCapabilityDto
{
    /// <summary>
    /// Nome completo del cmdlet (e.g., "Get-Mailbox", "Set-MailboxPermission").
    /// </summary>
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// True se il cmdlet è disponibile per l'utente corrente (RBAC permettendo).
    /// </summary>
    [JsonPropertyName("isAvailable")]
    public bool IsAvailable { get; set; }

    /// <summary>
    /// Lista dei parametri supportati dal cmdlet.
    /// Vuota se cmdlet non disponibile.
    /// </summary>
    [JsonPropertyName("parameters")]
    public List<string> Parameters { get; set; } = new();

    /// <summary>
    /// Motivo per cui il cmdlet non è disponibile (null se disponibile).
    /// Possibili valori: "Command not found", "Access denied", "Module not loaded".
    /// </summary>
    [JsonPropertyName("unavailableReason")]
    public string? UnavailableReason { get; set; }
}

/// <summary>
/// Feature-level capabilities derivate dai cmdlet disponibili.
/// Ogni flag indica se una specifica funzionalità è utilizzabile.
/// </summary>
/// <remarks>
/// <para>
/// <b>Feature Categories:</b>
/// <list type="table">
/// <listheader><term>Category</term><description>Capabilities</description></listheader>
/// <item><term>Mailbox Read</term><description>CanGetMailbox, CanGetMailboxStatistics</description></item>
/// <item><term>Mailbox Write</term><description>CanSetMailbox, CanSetArchive, CanSetLitigationHold</description></item>
/// <item><term>Permissions</term><description>CanGetMailboxPermission, CanAddMailboxPermission, CanRemoveMailboxPermission</description></item>
/// <item><term>Distribution Lists</term><description>CanGetDistributionGroup, CanGetDistributionGroupMember, CanAddDistributionGroupMember</description></item>
/// <item><term>Dynamic Groups</term><description>CanGetDynamicDistributionGroup, CanGetDynamicDistributionGroupMember</description></item>
/// <item><term>M365 Groups</term><description>CanGetUnifiedGroup, CanGetUnifiedGroupLinks</description></item>
/// </list>
/// </para>
/// </remarks>
public class FeatureCapabilitiesDto
{
    #region Mailbox Read Operations

    /// <summary>
    /// True se Get-Mailbox è disponibile (lettura mailbox).
    /// Richiesto per: lista mailbox, dettagli mailbox.
    /// </summary>
    [JsonPropertyName("canGetMailbox")]
    public bool CanGetMailbox { get; set; }

    /// <summary>
    /// True se Get-MailboxStatistics è disponibile (dimensioni, item count).
    /// Richiesto per: visualizzazione statistiche mailbox.
    /// </summary>
    [JsonPropertyName("canGetMailboxStatistics")]
    public bool CanGetMailboxStatistics { get; set; }

    /// <summary>
    /// True se Get-InboxRule è disponibile (regole inbox).
    /// Richiesto per: visualizzazione regole con forward detection.
    /// </summary>
    [JsonPropertyName("canGetInboxRule")]
    public bool CanGetInboxRule { get; set; }

    /// <summary>
    /// True se Get-MailboxAutoReplyConfiguration è disponibile (OOF).
    /// Richiesto per: visualizzazione stato fuori ufficio.
    /// </summary>
    [JsonPropertyName("canGetMailboxAutoReplyConfiguration")]
    public bool CanGetMailboxAutoReplyConfiguration { get; set; }

    #endregion

    #region Mailbox Write Operations

    /// <summary>
    /// True se Set-Mailbox è disponibile (modifica mailbox).
    /// Richiesto per: modifica proprietà mailbox.
    /// </summary>
    [JsonPropertyName("canSetMailbox")]
    public bool CanSetMailbox { get; set; }

    /// <summary>
    /// True se Enable-Mailbox -Archive / Set-Mailbox -ArchiveGuid è disponibile.
    /// Richiesto per: abilitazione/disabilitazione archivio.
    /// </summary>
    [JsonPropertyName("canSetArchive")]
    public bool CanSetArchive { get; set; }

    /// <summary>
    /// True se Set-Mailbox -LitigationHoldEnabled è disponibile.
    /// Richiesto per: abilitazione/disabilitazione litigation hold.
    /// </summary>
    [JsonPropertyName("canSetLitigationHold")]
    public bool CanSetLitigationHold { get; set; }

    /// <summary>
    /// True se Set-Mailbox -AuditEnabled è disponibile.
    /// Richiesto per: configurazione audit mailbox.
    /// </summary>
    [JsonPropertyName("canSetAudit")]
    public bool CanSetAudit { get; set; }

    #endregion

    #region Permission Operations

    /// <summary>
    /// True se Get-MailboxPermission è disponibile (FullAccess).
    /// Richiesto per: visualizzazione permessi FullAccess.
    /// </summary>
    [JsonPropertyName("canGetMailboxPermission")]
    public bool CanGetMailboxPermission { get; set; }

    /// <summary>
    /// True se Add-MailboxPermission è disponibile.
    /// Richiesto per: aggiunta permessi FullAccess.
    /// </summary>
    [JsonPropertyName("canAddMailboxPermission")]
    public bool CanAddMailboxPermission { get; set; }

    /// <summary>
    /// True se Remove-MailboxPermission è disponibile.
    /// Richiesto per: rimozione permessi FullAccess.
    /// </summary>
    [JsonPropertyName("canRemoveMailboxPermission")]
    public bool CanRemoveMailboxPermission { get; set; }

    /// <summary>
    /// True se Get-RecipientPermission è disponibile (SendAs).
    /// Richiesto per: visualizzazione permessi SendAs.
    /// </summary>
    [JsonPropertyName("canGetRecipientPermission")]
    public bool CanGetRecipientPermission { get; set; }

    /// <summary>
    /// True se Add-RecipientPermission è disponibile.
    /// Richiesto per: aggiunta permessi SendAs.
    /// </summary>
    [JsonPropertyName("canAddRecipientPermission")]
    public bool CanAddRecipientPermission { get; set; }

    /// <summary>
    /// True se Remove-RecipientPermission è disponibile.
    /// Richiesto per: rimozione permessi SendAs.
    /// </summary>
    [JsonPropertyName("canRemoveRecipientPermission")]
    public bool CanRemoveRecipientPermission { get; set; }

    #endregion

    #region Distribution List Operations

    /// <summary>
    /// True se Get-DistributionGroup è disponibile.
    /// Richiesto per: lista e dettagli gruppi di distribuzione.
    /// </summary>
    [JsonPropertyName("canGetDistributionGroup")]
    public bool CanGetDistributionGroup { get; set; }

    /// <summary>
    /// True se Set-DistributionGroup è disponibile.
    /// Richiesto per: modifica proprietà gruppi di distribuzione.
    /// </summary>
    [JsonPropertyName("canSetDistributionGroup")]
    public bool CanSetDistributionGroup { get; set; }

    /// <summary>
    /// True se Get-DistributionGroupMember è disponibile.
    /// Richiesto per: visualizzazione membri gruppo.
    /// </summary>
    [JsonPropertyName("canGetDistributionGroupMember")]
    public bool CanGetDistributionGroupMember { get; set; }

    /// <summary>
    /// True se Add-DistributionGroupMember è disponibile.
    /// Richiesto per: aggiunta membri a gruppo statico.
    /// </summary>
    [JsonPropertyName("canAddDistributionGroupMember")]
    public bool CanAddDistributionGroupMember { get; set; }

    /// <summary>
    /// True se Remove-DistributionGroupMember è disponibile.
    /// Richiesto per: rimozione membri da gruppo statico.
    /// </summary>
    [JsonPropertyName("canRemoveDistributionGroupMember")]
    public bool CanRemoveDistributionGroupMember { get; set; }

    #endregion

    #region Dynamic Distribution Group Operations

    /// <summary>
    /// True se Get-DynamicDistributionGroup è disponibile.
    /// Richiesto per: lista e dettagli DDG.
    /// </summary>
    [JsonPropertyName("canGetDynamicDistributionGroup")]
    public bool CanGetDynamicDistributionGroup { get; set; }

    /// <summary>
    /// True se Get-DynamicDistributionGroupMember è disponibile.
    /// Richiesto per: preview membri DDG (può essere lento).
    /// </summary>
    [JsonPropertyName("canGetDynamicDistributionGroupMember")]
    public bool CanGetDynamicDistributionGroupMember { get; set; }

    #endregion

    #region Microsoft 365 Group Operations

    /// <summary>
    /// True se Get-UnifiedGroup è disponibile.
    /// Richiesto per: lista M365 Groups.
    /// </summary>
    [JsonPropertyName("canGetUnifiedGroup")]
    public bool CanGetUnifiedGroup { get; set; }

    /// <summary>
    /// True se Get-UnifiedGroupLinks è disponibile.
    /// Richiesto per: visualizzazione membri/owner M365 Groups.
    /// </summary>
    [JsonPropertyName("canGetUnifiedGroupLinks")]
    public bool CanGetUnifiedGroupLinks { get; set; }

    #endregion
}

/// <summary>
/// Richiesta di detection capability.
/// </summary>
public class DetectCapabilitiesRequest
{
    /// <summary>
    /// Se true, forza il refresh delle capability anche se già presenti in cache.
    /// </summary>
    [JsonPropertyName("forceRefresh")]
    public bool ForceRefresh { get; set; }
}
