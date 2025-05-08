# thePOSHedgeFund

IYKYK

## Example

```powershell
$AKVEntraIdTenantId = '929ffa0f-08ae-4a91-bc05-2e99cee22249'
$AKVSubscriptionId = '4dd04a56-d486-4975-956e-34b85b62b688'
$AKVName = 'mmsmoa-2025'
$AKVUserIDSecretName = 'FrankQCUserID'
$AKVTokenSecretName = 'FrankQCToken'
$CSharpFilePath = Join-Path '.' 'algorithm-RSI-and-EMA.cs'

# Run the script:
. (Join-Path '.' 'Invoke-QCCryptoTrading.ps1') -AKVEntraIdTenantId $AKVEntraIdTenantId -AKVSubscriptionId $AKVSubscriptionId -AKVName $AKVName -AKVUserIDSecretName $AKVUserIDSecretName -AKVTokenSecretName $AKVTokenSecretName -CSharpFilePath $CSharpFilePath -DoNotCheckForModuleUpdates
```
