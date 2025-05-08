# thePOSHedgeFund

Here at POSHedgeFund, we take wealth building very seriously.

![an illustration of Frank Lesniak and Blake Cherry on a yacht in Monaco drinking sparkling wine. They are happy and appear to be celebrating something](./docs/celebration.png "Frank and Blake Living That Yacht Life in Monaco")

> [!TIP]
> "Trust us: we're professionals!"

## Disclaimer

The content of this presentation is for informational and educational purposes only and does not constitute financial, investment, or trading advice. The examples, scripts, and strategies discussed are illustrative and are not recommendations to buy, sell, or hold any securities or financial instruments. Past performance is not indicative of future results. Trading stocks and other financial instruments involves significant risk, including the risk of substantial financial loss. Attendees are solely responsible for their own investment decisions and should perform their own due diligence. This presentation does not establish any professional or advisory relationship between the presenter and the attendees. It is strongly recommended that attendees consult with a qualified financial advisor or professional before making any investment decisions.

All tools, code, and scripts provided during this session are offered "as is" without any warranties, express or implied, including but not limited to warranties of merchantability, fitness for a particular purpose, or non-infringement. The presenter and the conference organizers disclaim any liability for errors or omissions in the content or for any actions taken based on the information provided. In no event shall the presenter or organizers be liable for any direct, indirect, incidental, consequential, or punitive damages arising out of or related to the use of the information or materials provided. Attendees are responsible for ensuring that their use of any information or tools from this session complies with all applicable laws and regulations. All trademarks and third-party content are the property of their respective owners.

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
