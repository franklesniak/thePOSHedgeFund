using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.Drawing;
using QuantConnect;
using QuantConnect.Algorithm.Framework;
using QuantConnect.Algorithm.Framework.Selection;
using QuantConnect.Algorithm.Framework.Alphas;
using QuantConnect.Algorithm.Framework.Portfolio;
using QuantConnect.Algorithm.Framework.Portfolio.SignalExports;
using QuantConnect.Algorithm.Framework.Execution;
using QuantConnect.Algorithm.Framework.Risk;
using QuantConnect.Algorithm.Selection;
using QuantConnect.Api;
using QuantConnect.Parameters;
using QuantConnect.Benchmarks;
using QuantConnect.Brokerages;
using QuantConnect.Commands;
using QuantConnect.Configuration;
using QuantConnect.Util;
using QuantConnect.Interfaces;
using QuantConnect.Algorithm;
using QuantConnect.Indicators;
using QuantConnect.Data;
using QuantConnect.Data.Auxiliary;
using QuantConnect.Data.Consolidators;
using QuantConnect.Data.Custom;
using QuantConnect.Data.Custom.IconicTypes;
using QuantConnect.DataSource;
using QuantConnect.Data.Fundamental;
using QuantConnect.Data.Market;
using QuantConnect.Data.Shortable;
using QuantConnect.Data.UniverseSelection;
using QuantConnect.Notifications;
using QuantConnect.Orders;
using QuantConnect.Orders.Fees;
using QuantConnect.Orders.Fills;
using QuantConnect.Orders.OptionExercise;
using QuantConnect.Orders.Slippage;
using QuantConnect.Orders.TimeInForces;
using QuantConnect.Python;
using QuantConnect.Scheduling;
using QuantConnect.Securities;
using QuantConnect.Securities.Equity;
using QuantConnect.Securities.Future;
using QuantConnect.Securities.Option;
using QuantConnect.Securities.Positions;
using QuantConnect.Securities.Forex;
using QuantConnect.Securities.Crypto;
using QuantConnect.Securities.CryptoFuture;
using QuantConnect.Securities.IndexOption;
using QuantConnect.Securities.Interfaces;
using QuantConnect.Securities.Volatility;
using QuantConnect.Storage;
using QuantConnect.Statistics;
using QCAlgorithmFramework = QuantConnect.Algorithm.QCAlgorithm;
using QCAlgorithmFrameworkBridge = QuantConnect.Algorithm.QCAlgorithm;
using Calendar = QuantConnect.Data.Consolidators.Calendar;

public class BetterCryptoRsiEmaAlgo : QCAlgorithm
{
    // --- Parameters --------------------------------------------------------
    // List of available coinbase assets: https://www.quantconnect.com/docs/v2/writing-algorithms/datasets/coinapi/coinbase-crypto-price-data#11-Supported-Assets
    private readonly string[] _tickers = { "BTCUSD", "ETHUSD", "SOLUSD", "LTCUSD", "DOGEUSD" };
    private const int    RsiPeriod         = 14;
    private const int    EmaPeriod         = 50;
    private const int    SlopeLookbackBars = 5;          // 5 x 5-min = 25 min
    private const decimal PositionPct      = 0.05m;      // 5 % of equity per trade
    private const decimal StopLossPct      = -0.20m;     // -20 % unrealized
    private const decimal CircuitBreaker   = -0.50m;     // -50 % total
    private const decimal RSIOverbought    = 70.0m;      // RSI overbought threshold
    private const decimal RSIOverSold      = 30.0m;      // RSI oversold threshold
    private const decimal EMAUptrendPct = 0.0002m; // 0.02 % per bar (normalized slope)
    private const decimal EMADowntrendPct = -0.0002m; // -0.02 % per bar (normalized slope)

    // --- State -------------------------------------------------------------
    private readonly Dictionary<Symbol, RelativeStrengthIndex> _rsi = new();
    private readonly Dictionary<Symbol, ExponentialMovingAverage> _ema = new();
    private decimal _initialValue;
    private bool _haltTrading;
    private decimal _btcStartPrice;
    private bool _initialValueRecorded = false;
    private readonly Dictionary<Symbol, int> _liveTicksForSymbol = new();
    private const int MinLiveTicksForSlopeCalc = SlopeLookbackBars + 3; // e.g., 5 + 3 = 8 five-minute bars
    private int _onDataCallCountAfterWarmup = 0; // General counter for OnData calls post-warmup

    // -----------------------------------------------------------------------
    public override void Initialize()
    {
        SetStartDate(2024, 1, 1);
        SetEndDate(2025, 1, 15);          // demo window: 12½ months
        SetCash(10_000);
        SetBrokerageModel(BrokerageName.Coinbase, AccountType.Cash);

        // Add assets & indicators
        foreach (var ticker in _tickers)
        {
            var security = AddCrypto(ticker, Resolution.Minute, Market.Coinbase);
            var symbol = security.Symbol;
            _liveTicksForSymbol[symbol] = 0; // Initialize counter for each symbol

            var fiveMinConsolidator = ResolveConsolidator(symbol, TimeSpan.FromMinutes(5));

            var rsi = new RelativeStrengthIndex(RsiPeriod, MovingAverageType.Simple);
            var ema = new ExponentialMovingAverage(EmaPeriod);

            RegisterIndicator(symbol, rsi, fiveMinConsolidator);
            RegisterIndicator(symbol, ema, fiveMinConsolidator); // This ensures ema is updated by the consolidator

            _rsi[symbol] = rsi;
            _ema[symbol] = ema;
        }

        // Warm-up: enough bars for EMA plus slope look-back (dynamic)
        SetWarmUp(TimeSpan.FromMinutes((EmaPeriod + SlopeLookbackBars) * 5));

        // Benchmark seed (BTC first 5-min close)
        var btcSymbol = QuantConnect.Symbol.Create("BTCUSD", SecurityType.Crypto, Market.Coinbase);
        var hist = History<QuoteBar>(btcSymbol, 1, Resolution.Minute);
        try
        {
            if (hist != null && hist.Any())
            {
                _btcStartPrice = hist.First().Close;
                Log($"DEBUG: BTC Start Price for benchmark: {_btcStartPrice} at {hist.First().Time}");
            }
            else
            {
                Log("WARNING: Could not retrieve BTC history for benchmark start price during Initialize.");
                _btcStartPrice = 0m;
            }
        }
        catch (Exception ex)
        {
            Log($"ERROR: Getting BTC history in Initialize: {ex.Message}");
            _btcStartPrice = 0m;
        }
        Log("DEBUG: Initialize completed.");
    }

    // -----------------------------------------------------------------------
    public override void OnData(Slice data)
    {
        if (IsWarmingUp) return;

        _onDataCallCountAfterWarmup++;
        Log($"DEBUG: OnData Post-Warmup Call #{_onDataCallCountAfterWarmup} at {Time}. Slice contains keys: {string.Join(",", data.Keys.Select(k => k.Value))}");


        // Capture initial portfolio value once, right after warm-up completes
        if (!_initialValueRecorded)
        {
            if (Portfolio.TotalPortfolioValue > 0) // Ensure portfolio has a valid value
            {
                _initialValue = Portfolio.TotalPortfolioValue;
                _initialValueRecorded = true;
                Log($"INFO: Initial portfolio value recorded: {_initialValue:C} at {Time} (OnData post-warmup call #{_onDataCallCountAfterWarmup})");
            }
            else
            {
                // Portfolio value might be zero if cash hasn't settled or other reasons right at the start.
                // Depending on the brokerage model, this might need a slight delay or different handling.
                // For now, we'll log and it will try again on the next OnData if still 0.
                // Or, if 0 is unexpected, this indicates another issue.
                // However, with SetCash(10000), it should be > 0.
                Log($"WARNING: Portfolio.TotalPortfolioValue is 0 or less at {Time} (OnData post-warmup call #{_onDataCallCountAfterWarmup}) when attempting to record initial value. Will retry on next OnData.");
                return; // Don't proceed if portfolio value isn't sensible yet
            }
        }

        if (_haltTrading)
        {
            Log($"DEBUG: Trading is halted at {Time}.");
            return;
        }

        // --- Circuit breaker ------------------------------------------------
        // Ensure _initialValue has been set and is not zero before division
        if (_initialValueRecorded && _initialValue != 0m)
        {
            if ((Portfolio.TotalPortfolioValue / _initialValue - 1) < CircuitBreaker)
            {
                Log($"CRITICAL: Circuit breaker hit ({CircuitBreaker:P0}) at {Time}! Portfolio Value: {Portfolio.TotalPortfolioValue:C}, Initial Value: {_initialValue:C}. Halting new trades and liquidating.");
                _haltTrading = true;
                Liquidate();
                return;
            }
        }
        else if (!_initialValueRecorded && !IsWarmingUp)
        {
            // This case means we are past warmup, but _initialValue is still not set (e.g. if Portfolio.TotalPortfolioValue was 0 on first try)
            // We might want to prevent trading or log a warning, as circuit breaker cannot function yet.
            // For simplicity here, we let it pass, and it will try to record _initialValue again on the next OnData.
            // Debug("Circuit breaker cannot operate as initial portfolio value is not yet recorded.");
        }

        foreach (var symbol in _rsi.Keys) // Or _ema.Keys, or _liveTicksForSymbol.Keys
        {
            // Check if data for this symbol exists in the current slice
            if (!data.ContainsKey(symbol) || data[symbol] == null)
            {
                // Log($"TRACE: No data for {symbol} in current Slice at {Time}.");
                continue;
            }

            // Increment live tick counter only if we have data for this symbol
            _liveTicksForSymbol[symbol]++;

            var rsi = _rsi[symbol]; // Now 'symbol' is guaranteed to be a key
            var ema = _ema[symbol]; // Same here

            if (!rsi.IsReady || !ema.IsReady)
            {
                Log($"INFO: Indicators not ready for {symbol} at {Time}. RSI Ready: {rsi.IsReady} (Samples: {rsi.Samples}), EMA Ready: {ema.IsReady} (Samples: {ema.Samples}). Live ticks for symbol: {_liveTicksForSymbol[symbol]}");
                continue;
            }

            decimal calculatedEmaSlope = 0m;

            // Defer slope calculation until enough live 5-minute bars have been processed for this symbol
            if (_liveTicksForSymbol[symbol] < MinLiveTicksForSlopeCalc)
            {
                Log($"INFO: SLOPE DEFERRED for {symbol} at {Time}: Waiting for more live 5-min bars. Live 5-min ticks for symbol: {_liveTicksForSymbol[symbol]}/{MinLiveTicksForSlopeCalc}. EMA Samples: {ema.Samples}.");
                // calculatedEmaSlope remains 0m
            }
            else if (ema.Samples >= SlopeLookbackBars) // EMA must have at least processed SlopeLookbackBars number of 5-min bars
            {
                try
                {
                    if (SlopeLookbackBars > 1)
                    {
                        IndicatorDataPoint currentEmaDataPoint = ema.Current; // This is effectively ema[0]
                        IndicatorDataPoint pastEmaDataPoint = ema[SlopeLookbackBars - 1]; // Accessing ema[4]

                        // Corrected check: Focus on whether the IndicatorDataPoint objects themselves are null
                        if (currentEmaDataPoint == null || pastEmaDataPoint == null)
                        {
                            Log($"WARNING: SLOPE CALC for {symbol} at {Time}: EMA Current or Past IndicatorDataPoint is NULL. Current Exists: {currentEmaDataPoint != null}, Past (ema[{SlopeLookbackBars - 1}]) Exists: {pastEmaDataPoint != null}. EMA IsReady (overall): {ema.IsReady}, EMA Samples: {ema.Samples}, Live 5-min Ticks: {_liveTicksForSymbol[symbol]}. Slope set to 0.");
                            calculatedEmaSlope = 0m;
                        }
                        else
                        {
                            decimal currentEmaValue = currentEmaDataPoint.Value;
                            decimal pastEmaValue = pastEmaDataPoint.Value;

                            if (pastEmaValue != 0m)
                            {
                                calculatedEmaSlope = (currentEmaValue - pastEmaValue) / pastEmaValue / (SlopeLookbackBars - 1);
                                Log($"INFO: SLOPE CALC SUCCESS for {symbol} at {Time}: Slope={calculatedEmaSlope:0.######}. CurrentEMA: {currentEmaValue:F4}, PastEMA (ema[{SlopeLookbackBars - 1}]): {pastEmaValue:F4}. Live 5-min Ticks: {_liveTicksForSymbol[symbol]}");
                            }
                            else
                            {
                                Log($"WARNING: SLOPE CALC for {symbol} at {Time}: Past EMA value (ema[{SlopeLookbackBars - 1}]) is 0. Slope set to 0. Live 5-min Ticks: {_liveTicksForSymbol[symbol]}");
                                calculatedEmaSlope = 0m;
                            }
                        }
                    }
                    else if (SlopeLookbackBars == 1)
                    {
                        IndicatorDataPoint currentEmaDataPoint = ema.Current;
                        IndicatorDataPoint previousEmaDataPoint = ema.Previous; // This is ema[1]

                        // Corrected check
                        if (currentEmaDataPoint == null || previousEmaDataPoint == null)
                        {
                             Log($"WARNING: SLOPE CALC (1-bar lookback) for {symbol} at {Time}: EMA Current or Previous IndicatorDataPoint is NULL. Current Exists: {currentEmaDataPoint != null}, Previous Exists: {previousEmaDataPoint != null}. Slope set to 0. Live 5-min Ticks: {_liveTicksForSymbol[symbol]}");
                            calculatedEmaSlope = 0m;
                        }
                        else if (previousEmaDataPoint.Value != 0m)
                        {
                            calculatedEmaSlope = (currentEmaDataPoint.Value - previousEmaDataPoint.Value) / previousEmaDataPoint.Value;
                            Log($"INFO: SLOPE CALC SUCCESS (1-bar lookback) for {symbol} at {Time}: Slope={calculatedEmaSlope:0.######}. Live 5-min Ticks: {_liveTicksForSymbol[symbol]}");
                        }
                        else
                        {
                            Log($"WARNING: SLOPE CALC (1-bar lookback) for {symbol} at {Time}: EMA Previous value is 0. Slope set to 0. Live 5-min Ticks: {_liveTicksForSymbol[symbol]}");
                            calculatedEmaSlope = 0m;
                        }
                    }
                    // else: SlopeLookbackBars <= 0, slope remains 0m (shouldn't happen with const > 0)
                }
                catch (Exception ex)
                {
                    Error($"CRITICAL: Unhandled error in slope calculation logic for {symbol} at {Time}: {ex.Message}. StackTrace: {ex.StackTrace}. Slope set to 0.");
                    calculatedEmaSlope = 0m;
                }
            }
            else
            {
                 Log($"INFO: SLOPE DEFERRED for {symbol} at {Time}: EMA not yet enough samples ({ema.Samples}) for lookback ({SlopeLookbackBars}). Needs at least {SlopeLookbackBars} samples. Live 5-min Ticks: {_liveTicksForSymbol[symbol]}.");
            }

            var holdings = Portfolio[symbol].Quantity;

            bool upTrend = calculatedEmaSlope > EMAUptrendPct;
            bool downTrend = calculatedEmaSlope < EMADowntrendPct;

            Log($"DEBUG: TRADE CHECK for {symbol} at {Time}: Holdings: {holdings}, RSI: {rsi.Current.Value:F2}, CalcEMASlope: {calculatedEmaSlope:F6} (IsUpTrend: {upTrend}), TargetSlopeForUptrend: {EMAUptrendPct}, RSIOversold: {RSIOverSold}. Live 5-min Ticks: {_liveTicksForSymbol[symbol]}.");

            // ---------- Entries --------------------------------------------
            if (holdings == 0 && rsi.Current.Value < RSIOverSold && upTrend)
            {
                Log($"ACTION: ENTRY SIGNAL for {symbol} at {Time}: RSI ({rsi.Current.Value:F2}) < {RSIOverSold} AND EMA Slope ({calculatedEmaSlope:F6}) > {EMAUptrendPct}. Attempting SetHoldings({PositionPct:P0}).");
                SetHoldings(symbol, PositionPct, false, $"Entry: RSI < {RSIOverSold} & EMA UpTrend");
                // SetHoldings is asynchronous. Order fills will be handled by OnOrderEvent.
                // Debug($"{Time} BUY {symbol} | RSI {rsi.Current:0.##} EMA-slope {calculatedEmaSlope:0.#####}");
                continue; // Process next symbol
            }

            // ---------- Exits ----------------------------------------------
            if (holdings != 0)
            {
                bool exitedThisBar = false;
                // Stop-loss
                if (Portfolio[symbol].UnrealizedProfitPercent <= StopLossPct)
                {
                    Log($"ACTION: EXIT SIGNAL (Stop-Loss) for {symbol} at {Time}: UnrealizedProfitPercent ({Portfolio[symbol].UnrealizedProfitPercent:P2}) <= {StopLossPct:P2}. Liquidating.");
                    Liquidate(symbol: symbol, tag: "stop-loss");
                    exitedThisBar = true;
                }

                // RSI profit-take
                if (!exitedThisBar && rsi.Current.Value > RSIOverbought)
                {
                    Log($"ACTION: EXIT SIGNAL (RSI Profit Take) for {symbol} at {Time}: RSI ({rsi.Current.Value:F2}) > {RSIOverbought}. Liquidating.");
                    Liquidate(symbol: symbol, tag: $"RSI > {RSIOverbought}");
                    // Debug($"{Time} SELL {symbol} | RSI {rsi.Current:0.##} EMA-slope {calculatedEmaSlope:0.#####}");
                    // You might want a continue here if this exit should prevent the demo exit check
                    // continue;
                    exitedThisBar = true;
                }

                // Demo-specific exit: Close after 30 mins or 1.5% profit
                if (!exitedThisBar)
                {
                    DateTime? lastFillTimeForSymbol = null;
                    var relevantTickets = Transactions.GetOrderTickets(o => o.Symbol == symbol && (o.Status == OrderStatus.Filled || o.Status == OrderStatus.PartiallyFilled));
                    if (relevantTickets.Any())
                    {
                        lastFillTimeForSymbol = relevantTickets.SelectMany(t => t.OrderEvents)
                                                              .Where(oe => oe.Status == OrderStatus.Filled || (oe.Status == OrderStatus.PartiallyFilled && oe.FillQuantity != 0))
                                                              .Select(oe => oe.UtcTime) // Use UtcTime
                                                              .DefaultIfEmpty(DateTime.MinValue)
                                                              .Max();
                        if (lastFillTimeForSymbol == DateTime.MinValue) lastFillTimeForSymbol = null; // Ensure it's null if no valid fills
                    }

                    bool timeExitCondition = false;
                    if (lastFillTimeForSymbol.HasValue)
                    {
                        // Algorithm.Time is also in UTC when running in cloud.
                        timeExitCondition = (Time - lastFillTimeForSymbol.Value) > TimeSpan.FromMinutes(30);
                    }

                    if (timeExitCondition || Portfolio[symbol].UnrealizedProfitPercent > 0.015m)
                    {
                        Log($"ACTION: EXIT SIGNAL (Demo Time/Profit) for {symbol} at {Time}. Time Exit: {timeExitCondition}, Profit Exit: {Portfolio[symbol].UnrealizedProfitPercent > 0.015m}. Last Fill: {lastFillTimeForSymbol?.ToString("o") ?? "N/A"}. Liquidating.");
                        Liquidate(symbol: symbol, tag: "Demo Time/Profit Exit");
                        exitedThisBar = true;
                    }
                }
                if (exitedThisBar) continue; // Process next symbol
            }
        }
        // Log($"DEBUG: OnData Post-Warmup Call #{_onDataCallCountAfterWarmup} COMPLETED at {Time}.");
    }

    public override void OnOrderEvent(OrderEvent orderEvent)
    {
        // Log every order event with details
        Log($"INFO: ORDER EVENT: Symbol: {orderEvent.Symbol.Value}, DateTime: {orderEvent.UtcTime.ToString("o")}, Status: {orderEvent.Status}, OrderId: {orderEvent.OrderId}, FillQuantity: {orderEvent.FillQuantity}, FillPrice: {orderEvent.FillPrice:F4}, Direction: {orderEvent.Direction}, Message: {orderEvent.Message}");

        // Example: Log when an order is filled
        if (orderEvent.Status == OrderStatus.Filled || (orderEvent.Status == OrderStatus.PartiallyFilled && orderEvent.FillQuantity != 0))
        {
            Log($"INFO: TRADE EXECUTED: {orderEvent.Direction} {orderEvent.AbsoluteFillQuantity} of {orderEvent.Symbol.Value} at {orderEvent.FillPrice:F4}. OrderId: {orderEvent.OrderId}");
        }
        base.OnOrderEvent(orderEvent);
    }

    // -----------------------------------------------------------------------
    public override void OnEndOfAlgorithm()
    {
        Log($"INFO: OnEndOfAlgorithm called at {Time}. Final Portfolio Value: {Portfolio.TotalPortfolioValue:C}");
        if (_initialValueRecorded && _initialValue > 0)
        {
            var perf = (Portfolio.TotalPortfolioValue / _initialValue - 1) * 100;
            Log($"FINAL: Strategy P&L: {perf:0.##}% (Initial: {_initialValue:C}, Final: {Portfolio.TotalPortfolioValue:C})");
        }
        else if (!_initialValueRecorded)
        {
            Log("FINAL: Strategy P&L: N/A (Initial portfolio value was not recorded).");
        }
        else
        {
            Log("FINAL: Strategy P&L: N/A (Initial portfolio value was zero).");
        }

        if (_btcStartPrice > 0 && Securities.ContainsKey("BTCUSD"))
        {
            var btcSecurity = Securities["BTCUSD"];
            if (btcSecurity.Price > 0)
            {
                var btcEnd = btcSecurity.Price;
                var hodl = (btcEnd / _btcStartPrice - 1) * 100;
                Log($"FINAL: BTC Buy-and-Hold P&L (from algo start): {hodl:0.##}% (Start BTC Price: {_btcStartPrice:F2}, End BTC Price: {btcEnd:F2})");
                if (_initialValueRecorded && _initialValue > 0)
                {
                    var perf = (Portfolio.TotalPortfolioValue / _initialValue - 1) * 100;
                    Log($"FINAL: Strategy vs BTC HODL: Δ = {(perf - hodl):0.##}%");
                }
            }
            else
            {
                Log("FINAL: BTC Buy-and-Hold P&L: N/A (BTC end price is zero or unavailable).");
            }
        }
        else
        {
            Log("FINAL: BTC Buy-and-Hold P&L: N/A (BTC start price was zero or BTCUSD security not found).");
        }
        Log("----- Algorithm Execution Finished -----");
    }
}
