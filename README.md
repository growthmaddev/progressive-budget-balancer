# Enhanced Design Document: Progressive Budget Balancer for Google Ads

## 1. Purpose and Overview

The Progressive Budget Balancer is an advanced script designed to optimize Google Ads budgets through intelligent allocation across campaigns of varying objective types. The script makes data-driven decisions based on multiple performance factors, bid strategy-specific goals, and temporal performance patterns to maximize campaign effectiveness within budget constraints.

### Core Objectives:

- Progressively balance budgets based on multifaceted performance analysis
- Support all Google Ads campaign objective types and bid strategies
- Maintain monthly budget targets with daily optimization
- Account for day-of-week performance patterns with adaptive analysis
- Support both primary and specific conversion action analysis
- Enable shared budget handling via API-based relationship detection
- Provide detailed insights into budget adjustment decisions
- Make incremental changes that respect campaign-specific performance patterns
- Handle edge cases including new campaigns and limited data scenarios

## 2. System Architecture

### 2.1 Execution Flow

1. **Initialization** - Set up configuration and prepare execution environment
2. **Monthly Budget Calculation** - Determine daily budget targets based on monthly goals
3. **Day-of-Week Analysis** - Analyze historical day-of-week performance patterns with adaptive lookback
4. **Data Collection** - Gather 90-day, 21-day and recent performance data
5. **Performance Analysis** - Calculate efficiency metrics and trends by bid strategy
6. **Budget Adjustment Calculation** - Determine optimal budget changes
7. **Zero-Sum Balancing** - Scale adjustments to maintain budget targets
8. **Budget Application** - Apply changes to Google Ads campaigns
9. **Reporting** - Log detailed summary of actions and reasons
10. **Error Recovery** - Handle exceptions and ensure script completion

### 2.2 Component Overview

- **Configuration Module** - User-defined settings
- **Budget Planning Module** - Monthly and daily budget management
- **Temporal Analysis Module** - Adaptive day-of-week performance patterns
- **Data Collection Module** - Strategy-specific metric retrieval
- **Analysis Engine** - Multifaceted performance evaluation
- **Adjustment Engine** - Budget decision algorithms
- **Execution Module** - Budget change implementation
- **Reporting Module** - Logging and summary generation
- **Error Handling Module** - Exception management and recovery
- **History Module** - Store and retrieve historical adjustments

## 3. Data Structures

### 3.1 Campaign Data Object

Each campaign is represented by an object containing:

```javascript
{
  campaign: Object, // Google Ads campaign object
  name: String, // Campaign name
  bidStrategy: String, // Optimization strategy
  isSharedBudget: Boolean, // Shared budget flag
  sharedBudgetId: String, // Identifier for shared budget group
  currentDailyBudget: Number, // Current budget amount
  recommendedDailyBudget: Number, // Day-adjusted budget
  
  // Performance metrics
  cost: Number, // Historical cost
  objectiveMetrics: {
    byPeriod: {
      ninetyDay: Object, // 90-day metrics
      twentyOneDay: Object, // 21-day metrics
      sevenDay: Object, // 7-day metrics
      dayOfWeek: Object // Metrics by day of week
    },
    volume: Number, // Objective-specific volume metric
    performance: Number, // Objective-specific performance metric
    target: Number // Objective-specific target value
  },
  
  // Conversion metrics
  conversions: {
    primary: Number, // Default conversion count
    specific: Number, // Specific conversion action count
    value: Number // Conversion value
  },
  
  // Impression metrics
  impressions: Number, // Impression count
  impressionShare: Number, // Total impression share
  budgetImpressionShareLost: Number, // Impression share lost to budget
  
  // Data sufficiency flags
  hasMinimumData: Boolean, // Whether campaign has sufficient data for full analysis
  daysSinceCreation: Number, // Days since campaign was created
  
  // Scoring factors
  scores: {
    efficiency: Number, // Volume/budget ratio
    recency: Number, // Recent vs historical performance
    spendUpside: Number, // Based on impression share loss
    dayOfWeek: {
      score: Number, // Day-specific performance
      confidence: Boolean, // Whether we have statistical confidence
      daysAnalyzed: Number, // Total days analyzed in lookback
      conversionsAnalyzed: Number // Total conversions used in analysis
    },
    combined: Number // Overall score
  },
  
  // Budget calculations
  budgetPercentage: Number, // % of total budget
  volumePercentage: Number, // % of total objective volume
  gap: Number, // Difference between volume % and budget %
  
  // Budget decision results
  adjustmentFactors: Object, // Detailed adjustment factors
  newBudget: Number, // Calculated new budget
  adjustmentReason: String, // Human-readable explanation
  
  // Error states
  hasError: Boolean, // Whether any errors occurred during processing
  errorDetails: String, // Description of any errors
  conversionSource: String, // Source of conversion data
  conversionAction: String // Specific conversion action used
}
```

### 3.2 Bid Strategy Performance Metrics

Strategy-specific performance metrics:

```javascript
{
  'MAXIMIZE_CLICKS': {
    volumeMetric: 'clicks',
    performanceMetric: 'ctr',
    lowerIsBetter: false,
    targetMetric: null
  },
  'MAXIMIZE_CONVERSIONS': {
    volumeMetric: 'conversions',
    performanceMetric: 'convRate',
    lowerIsBetter: false,
    targetMetric: null
  },
  'TARGET_CPA': {
    volumeMetric: 'conversions',
    performanceMetric: 'cpa',
    lowerIsBetter: true,
    targetMetric: 'targetCpa'
  },
  'TARGET_ROAS': {
    volumeMetric: 'conversionValue',
    performanceMetric: 'roas',
    lowerIsBetter: false,
    targetMetric: 'targetRoas'
  },
  'MAXIMIZE_CONVERSION_VALUE': {
    volumeMetric: 'conversionValue',
    performanceMetric: 'valuePerCost',
    lowerIsBetter: false,
    targetMetric: null
  },
  'TARGET_IMPRESSION_SHARE': {
    volumeMetric: 'impressions',
    performanceMetric: 'impressionShare',
    lowerIsBetter: false,
    targetMetric: 'targetImpressionShare'
  },
  'MANUAL_CPC': {
    volumeMetric: 'conversions',
    performanceMetric: 'roi',
    lowerIsBetter: false,
    targetMetric: null
  }
}
```

### 3.3 Day-of-Week Performance Data

```javascript
{
  // For each campaign
  campaignId: {
    // By day of week (0 = Sunday, 6 = Saturday)
    0: { volume: Number, performance: Number, cost: Number, conversions: Number },
    1: { volume: Number, performance: Number, cost: Number, conversions: Number },
    2: { volume: Number, performance: Number, cost: Number, conversions: Number },
    3: { volume: Number, performance: Number, cost: Number, conversions: Number },
    4: { volume: Number, performance: Number, cost: Number, conversions: Number },
    5: { volume: Number, performance: Number, cost: Number, conversions: Number },
    6: { volume: Number, performance: Number, cost: Number, conversions: Number },
    // Indexed performance (1.0 = average)
    dayIndex: {
      0: Number, 1: Number, 2: Number, 3: Number, 4: Number, 5: Number, 6: Number
    },
    // Analysis metadata
    analysisMetadata: {
      confidence: Boolean, // Whether we have statistical confidence
      totalDaysAnalyzed: Number, // Total days used in analysis
      totalConversions: Number, // Total conversions analyzed
      blockCount: Number // Number of 7-day blocks used
    }
  }
}
```

### 3.4 Historical Adjustment Record

```javascript
{
  date: String, // ISO date of adjustment
  campaign: {
    id: String,
    name: String,
    bidStrategy: String
  },
  budgets: {
    before: Number,
    after: Number,
    percentChange: Number
  },
  factors: {
    efficiency: Number,
    recency: Number,
    spendUpside: Number,
    dayOfWeek: Number
  },
  metrics: {
    // Relevant metrics at time of adjustment
    cost: Number,
    volume: Number,
    performance: Number
  },
  reason: String // Explanation of adjustment
}
```

## 4. Required Functions

### 4.1 Main Execution Function

```javascript
function main() {
  try {
    // Initialize configuration and logging
    const config = initializeConfig();
    Logger.log("Starting Progressive Budget Balancer execution");
    
    // Calculate monthly budget pacing
    const budgetData = calculateMonthlyBudgetPacing(config);
    logBudgetPacingStatus(budgetData);
    
    // Get all campaigns for processing
    const campaigns = getCampaignsForProcessing(config);
    Logger.log(`Found ${campaigns.length} campaigns for processing`);
    
    // Group campaigns by shared budget
    const { sharedBudgetGroups, individualCampaigns } = groupCampaignsByBudget(campaigns);
    
    // Get current day of week for day-specific performance
    const currentDate = new Date();
    const currentDayOfWeek = currentDate.getDay();
    
    // Process individual campaigns
    const processedIndividualCampaigns = processCampaignBudgets(
      individualCampaigns, 
      currentDayOfWeek,
      budgetData,
      config
    );
    
    // Process shared budget campaigns
    const processedSharedBudgets = processSharedBudgets(
      sharedBudgetGroups, 
      currentDayOfWeek, 
      budgetData,
      config
    );
    
    // Combine all processed campaigns for reporting
    const allProcessedCampaigns = [
      ...processedIndividualCampaigns,
      ...processedSharedBudgets.flatMap(group => group.campaigns)
    ];
    
    // Apply budget changes
    applyBudgetChanges(allProcessedCampaigns, processedSharedBudgets);
    
    // Store adjustment history
    storeAdjustmentHistory(allProcessedCampaigns);
    
    // Generate summary report
    generateSummaryReport(allProcessedCampaigns, budgetData);
    
    Logger.log("Progressive Budget Balancer execution completed successfully");
  } catch (error) {
    Logger.log(`ERROR: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    
    // Try to send email alert if critical error
    try {
      sendErrorAlert(error);
    } catch (emailError) {
      Logger.log(`Failed to send error alert: ${emailError.message}`);
    }
  }
}
```

### 4.2 Budget Management Functions

```javascript
function calculateMonthlyBudgetPacing(config) {
  try {
    // Get current date info
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth();
    
    // Calculate days in current month
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const today = now.getDate();
    const daysRemaining = daysInMonth - today + 1;
    const daysElapsed = today - 1;
    
    // Calculate month-to-date spend
    const mtdSpend = getMonthToDateSpend();
    
    // Calculate remaining budget
    const monthlyBudget = config.MONTHLY_BUDGET;
    const remainingBudget = monthlyBudget - mtdSpend;
    
    // Calculate daily budget target for remaining days
    const dailyBudgetTarget = remainingBudget / daysRemaining;
    
    // Calculate pacing metrics
    const idealSpendToDate = (monthlyBudget / daysInMonth) * daysElapsed;
    const pacingDifference = mtdSpend - idealSpendToDate;
    const pacingStatus = pacingDifference > 0 ? 'AHEAD' : 'BEHIND';
    const pacingPercentage = (mtdSpend / idealSpendToDate) * 100;
    
    // Projected spend by month end
    const projectedSpend = mtdSpend + (dailyBudgetTarget * daysRemaining);
    
    return {
      monthlyBudget: monthlyBudget,
      mtdSpend: mtdSpend,
      daysInMonth: daysInMonth,
      daysRemaining: daysRemaining,
      daysElapsed: daysElapsed,
      remainingBudget: remainingBudget,
      dailyBudgetTarget: dailyBudgetTarget,
      pacingDifference: pacingDifference,
      pacingStatus: pacingStatus,
      pacingPercentage: pacingPercentage,
      projectedSpend: projectedSpend,
      isEndOfMonth: daysRemaining <= 7
    };
  } catch (error) {
    Logger.log(`Error in calculateMonthlyBudgetPacing: ${error.message}`);
    throw error;
  }
}

function getDayOfWeekPerformancePattern(campaign, initialDateRange, config) {
  try {
    // Configuration
    const MIN_CONVERSIONS_FOR_CONFIDENCE = config.DOW_MIN_CONVERSIONS_FOR_CONFIDENCE;
    const MAX_LOOKBACK_BLOCKS = config.DOW_MAX_LOOKBACK_BLOCKS;
    const BLOCK_SIZE = config.DOW_BLOCK_SIZE;
    
    let totalLookbackDays = initialDateRange; // Start with initial lookback (e.g., 90 days)
    let conversionsByDay = initializeDayOfWeekData();
    let totalConversions = 0;
    let currentBlock = 0;
    let prevBlockConversions = -1; // Track if new blocks add conversions
    
    // Loop until we have enough conversions or hit limits
    while (totalConversions < MIN_CONVERSIONS_FOR_CONFIDENCE && 
           currentBlock < MAX_LOOKBACK_BLOCKS && 
           (currentBlock === 0 || totalConversions > prevBlockConversions)) {
      
      // Extend lookback range if not first iteration
      if (currentBlock > 0) {
        prevBlockConversions = totalConversions;
        totalLookbackDays += BLOCK_SIZE;
      }
      
      // Get data for current lookback period
      const dateRange = getDateRangeForLookback(totalLookbackDays);
      const conversionData = getConversionDataByDayOfWeek(campaign, dateRange, config);
      
      // Update day of week data with new block
      conversionsByDay = mergeConversionData(conversionsByDay, conversionData);
      totalConversions = calculateTotalConversions(conversionsByDay);
      
      // Increment block counter
      currentBlock++;
    }
    
    // If we don't have sufficient data, return neutral factors
    if (totalConversions < MIN_CONVERSIONS_FOR_CONFIDENCE) {
      return {
        dayIndex: {0: 1.0, 1: 1.0, 2: 1.0, 3: 1.0, 4: 1.0, 5: 1.0, 6: 1.0},
        hasConfidence: false,
        totalDaysAnalyzed: totalLookbackDays,
        totalConversions: totalConversions,
        blockCount: currentBlock
      };
    }
    
    // Calculate day of week index from collected data
    const dayIndex = calculateDayPerformanceIndex(conversionsByDay);
    
    return {
      dayIndex: dayIndex,
      hasConfidence: true,
      totalDaysAnalyzed: totalLookbackDays,
      totalConversions: totalConversions,
      blockCount: currentBlock
    };
  } catch (error) {
    Logger.log(`Error in getDayOfWeekPerformancePattern for campaign ${campaign.getName()}: ${error.message}`);
    // Return neutral factors if error occurs
    return {
      dayIndex: {0: 1.0, 1: 1.0, 2: 1.0, 3: 1.0, 4: 1.0, 5: 1.0, 6: 1.0},
      hasConfidence: false,
      totalDaysAnalyzed: 0,
      totalConversions: 0,
      blockCount: 0,
      error: error.message
    };
  }
}

// Helper functions for day-of-week analysis

function initializeDayOfWeekData() {
  // Create empty data structure for day of week data
  const days = [0, 1, 2, 3, 4, 5, 6]; // Sunday to Saturday
  const result = {};
  
  days.forEach(day => {
    result[day] = {
      volume: 0,
      performance: 0,
      cost: 0,
      conversions: 0,
      days: 0 // Count of days analyzed for this day of week
    };
  });
  
  return result;
}

function getDateRangeForLookback(days) {
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - days);
  
  return {
    startDate: startDate.toISOString().split('T')[0],
    endDate: endDate.toISOString().split('T')[0]
  };
}

function mergeConversionData(existingData, newData) {
  // Merge new conversion data into existing structure
  const result = JSON.parse(JSON.stringify(existingData)); // Deep copy
  
  Object.keys(newData).forEach(day => {
    if (result[day]) {
      result[day].volume += newData[day].volume || 0;
      result[day].cost += newData[day].cost || 0;
      result[day].conversions += newData[day].conversions || 0;
      result[day].days += newData[day].days || 1;
      
      // Update performance metrics based on new totals
      if (result[day].volume > 0 && result[day].cost > 0) {
        result[day].performance = result[day].conversions / result[day].volume;
      }
    }
  });
  
  return result;
}

function calculateTotalConversions(dayOfWeekData) {
  // Sum up all conversions across days of week
  return Object.values(dayOfWeekData).reduce((sum, day) => sum + (day.conversions || 0), 0);
}

function calculateDayPerformanceIndex(dayOfWeekData) {
  // Calculate performance indices for each day compared to average
  const days = Object.keys(dayOfWeekData);
  
  // Calculate average daily performance
  let totalPerformance = 0;
  let validDayCount = 0;
  
  days.forEach(day => {
    if (dayOfWeekData[day].conversions > 0) {
      totalPerformance += dayOfWeekData[day].performance;
      validDayCount++;
    }
  });
  
  const avgPerformance = validDayCount > 0 ? totalPerformance / validDayCount : 0;
  
  // Calculate index for each day (relative to average)
  const dayIndex = {};
  
  days.forEach(day => {
    if (avgPerformance > 0 && dayOfWeekData[day].conversions > 0) {
      dayIndex[day] = dayOfWeekData[day].performance / avgPerformance;
    } else {
      dayIndex[day] = 1.0; // Neutral if no data
    }
    
    // Limit the range of adjustment factors (e.g., 0.8 to 1.2)
    dayIndex[day] = Math.max(0.8, Math.min(1.2, dayIndex[day]));
  });
  
  return dayIndex;
}
```

### 4.3 Data Collection Functions

```javascript
function getStrategySpecificMetrics(campaign, dateRange, bidStrategy, config) {
  try {
    // Find the appropriate metrics for this bid strategy
    const strategyMetrics = STRATEGY_METRICS[bidStrategy];
    if (!strategyMetrics) {
      // Fallback to a default strategy if unknown
      Logger.log(`Unknown bid strategy ${bidStrategy} for campaign ${campaign.getName()}, using MANUAL_CPC metrics`);
      return getStrategySpecificMetrics(campaign, dateRange, 'MANUAL_CPC', config);
    }
    
    const volumeMetric = strategyMetrics.volumeMetric;
    const performanceMetric = strategyMetrics.performanceMetric;
    const targetMetric = strategyMetrics.targetMetric;
    
    // Prepare the query based on metrics needed
    const query = buildMetricsQuery(campaign, dateRange, strategyMetrics);
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    if (!rows.hasNext()) {
      // No data available
      return {
        volume: 0,
        performance: 0,
        target: null,
        cost: 0,
        hasData: false
      };
    }
    
    // Process the data
    let row = rows.next();
    let metrics = {
      volume: getMetricValue(row, volumeMetric),
      performance: getMetricValue(row, performanceMetric),
      target: targetMetric ? campaign[targetMetric] : null,
      cost: getMetricValue(row, 'cost'),
      hasData: true
    };
    
    // For strategies with explicit targets, calculate performance against target
    if (targetMetric && metrics.target) {
      if (strategyMetrics.lowerIsBetter) {
        metrics.performanceScore = metrics.target / metrics.performance;
      } else {
        metrics.performanceScore = metrics.performance / metrics.target;
      }
    } else {
      metrics.performanceScore = 1; // Neutral score for strategies without targets
    }
    
    return metrics;
  } catch (error) {
    Logger.log(`Error getting metrics for campaign ${campaign.getName()}: ${error.message}`);
    
    // Return neutral/empty metrics on error
    return {
      volume: 0,
      performance: 0,
      target: null,
      cost: 0,
      hasData: false,
      error: error.message
    };
  }
}

function getConversionData(campaign, dateRange, config) {
  try {
    const conversionAction = resolveConversionAction(campaign);
    
    // Build query based on whether using alternate conversions
    const query = buildConversionQuery(campaign, dateRange, conversionAction);
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    if (!rows.hasNext()) {
      return {
        conversions: 0,
        conversionValue: 0,
        hasData: false,
        conversionSource: conversionAction ? 'alternate' : 'default'
      };
    }
    
    const row = rows.next();
    return {
      conversions: getMetricValue(row, conversionAction ? 'AllConv' : 'Conversions'),
      conversionValue: getMetricValue(row, conversionAction ? 'AllConvValue' : 'ConversionValue'),
      hasData: true,
      conversionSource: conversionAction ? 'alternate' : 'default',
      conversionAction: conversionAction
    };
  } catch (error) {
    Logger.log(`Error getting conversion data for campaign ${campaign.getName()}: ${error.message}`);
    return {
      conversions: 0,
      conversionValue: 0,
      hasData: false,
      error: error.message
    };
  }
}

function getConversionDataByDayOfWeek(campaign, dateRange, config) {
  try {
    // Query conversions with day breakdown
    const query = buildDayOfWeekQuery(campaign, dateRange);
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Initialize result structure
    const result = initializeDayOfWeekData();
    
    // Process each row (each day's data)
    while (rows.hasNext()) {
      const row = rows.next();
      const date = new Date(row['Date']);
      const dayOfWeek = date.getDay(); // 0 = Sunday, 6 = Saturday
      
      // Get metrics for this day
      const conversions = getMetricValue(row, 'conversions');
      const cost = getMetricValue(row, 'cost');
      const volume = getMetricValue(row, 'clicks'); // Or other volume metric based on strategy
      
      // Add to appropriate day of week
      result[dayOfWeek].conversions += conversions;
      result[dayOfWeek].cost += cost;
      result[dayOfWeek].volume += volume;
      result[dayOfWeek].days += 1;
      
      // Calculate performance metric (e.g., conversion rate)
      if (volume > 0) {
        result[dayOfWeek].performance = result[dayOfWeek].conversions / result[dayOfWeek].volume;
      }
    }
    
    return result;
  } catch (error) {
    Logger.log(`Error getting day-of-week data for campaign ${campaign.getName()}: ${error.message}`);
    
    // Return empty structure on error
    return initializeDayOfWeekData();
  }
}

// Helper functions for data collection

function buildMetricsQuery(campaign, dateRange, strategyMetrics) {
  // Build appropriate AWQL query based on strategy metrics
  // Implementation details would go here
}

function buildConversionQuery(campaign, dateRange, conversionAction) {
  // Build AWQL query for conversion data
  // Implementation details would go here
}

function buildDayOfWeekQuery(campaign, dateRange) {
  // Build AWQL query with day breakdown
  // Implementation details would go here
}

function getMetricValue(row, metricName) {
  // Safely extract numeric value from report row
  try {
    const value = row[metricName];
    return parseFloat(value.replace(/,/g, '')) || 0;
  } catch (e) {
    return 0;
  }
}

function resolveConversionAction(campaign) {
  try {
    // If not using alternate conversions, return null to use default
    if (!CONFIG.CONVERSION_SETTINGS.USE_ALTERNATE_CONVERSIONS) {
      return null;
    }
    
    // Get campaign ID
    const campaignId = campaign.getId();
    
    // If using campaign-specific mappings, try to get specific action
    if (CONFIG.CONVERSION_SETTINGS.USE_CAMPAIGN_SPECIFIC_MAPPINGS) {
      const specificAction = CONFIG.CONVERSION_SETTINGS.CAMPAIGN_MAPPINGS[campaignId];
      if (specificAction) {
        return specificAction;
      }
    }
    
    // Use master alternate conversion as fallback
    return CONFIG.CONVERSION_SETTINGS.MASTER_ALTERNATE_CONVERSION;
  } catch (error) {
    Logger.log(`Error resolving conversion action for campaign ${campaign.getName()}: ${error.message}`);
    return CONFIG.CONVERSION_SETTINGS.MASTER_ALTERNATE_CONVERSION;
  }
}
```

### 4.4 Analysis Functions

```javascript
function calculateEfficiencyScore(campaign, campaignCollection) {
  try {
    // Calculate campaign's percentage of total objective volume
    const totalVolume = campaignCollection.reduce((sum, c) => sum + c.objectiveMetrics.volume, 0);
    const volumePercentage = totalVolume > 0 ? 
      (campaign.objectiveMetrics.volume / totalVolume) * 100 : 0;
    
    // Calculate campaign's percentage of total budget
    const totalBudget = campaignCollection.reduce((sum, c) => sum + c.currentDailyBudget, 0);
    const budgetPercentage = totalBudget > 0 ?
      (campaign.currentDailyBudget / totalBudget) * 100 : 0;
    
    // Store these percentages for reporting
    campaign.volumePercentage = volumePercentage;
    campaign.budgetPercentage = budgetPercentage;
    
    // Calculate efficiency (volume % / budget %)
    let efficiencyScore;
    if (budgetPercentage > 0) {
      efficiencyScore = volumePercentage / budgetPercentage;
    } else {
      efficiencyScore = 1; // Neutral score if no budget
    }
    
    // Cap extreme values
    efficiencyScore = Math.max(0.5, Math.min(2.0, efficiencyScore));
    
    // Calculate the gap between volume and budget percentages for reporting
    campaign.gap = volumePercentage - budgetPercentage;
    
    return efficiencyScore;
  } catch (error) {
    Logger.log(`Error calculating efficiency score for campaign ${campaign.name}: ${error.message}`);
    return 1; // Return neutral score on error
  }
}

function calculateRecencyScore(campaign, recentPeriod, historicalPeriod) {
  try {
    // Check if we have valid data for both periods
    if (!recentPeriod.hasData || !historicalPeriod.hasData) {
      return 1; // Neutral score if missing data
    }
    
    // Get performance metrics for both periods
    const recentPerformance = recentPeriod.performance;
    const historicalPerformance = historicalPeriod.performance;
    
    // Check for zero values to avoid division errors
    if (historicalPerformance === 0 || recentPerformance === 0) {
      return 1; // Neutral score if zero performance
    }
    
    // Get strategy to determine if lower is better
    const bidStrategy = campaign.bidStrategy;
    const strategyMetrics = STRATEGY_METRICS[bidStrategy];
    
    let recencyScore;
    
    if (strategyMetrics.lowerIsBetter) {
      // For metrics like CPA where lower is better (inversely proportional)
      recencyScore = historicalPerformance / recentPerformance;
    } else {
      // For metrics like CTR, conv rate where higher is better
      recencyScore = recentPerformance / historicalPerformance;
    }
    
    // Cap extreme values
    recencyScore = Math.max(0.5, Math.min(2.0, recencyScore));
    
    return recencyScore;
  } catch (error) {
    Logger.log(`Error calculating recency score for campaign ${campaign.name}: ${error.message}`);
    return 1; // Return neutral score on error
  }
}

function calculateSpendUpsideScore(campaign, config) {
  try {
    // Get impression share lost due to budget
    const budgetImpressionShareLost = campaign.budgetImpressionShareLost || 0;
    
    // Check if this is above our threshold
    const threshold = config.MAX_IMPRESSION_SHARE_LOST_TO_BUDGET;
    
    if (budgetImpressionShareLost > threshold) {
      // Scaling factor based on how much impression share is being lost
      // Higher loss = higher score (more budget needed)
      const lostFactor = budgetImpressionShareLost / 100;
      
      // Scale to desired range (e.g., 1.0 to 1.5)
      const spendUpsideScore = 1 + (lostFactor * 0.5);
      
      // Cap at maximum value
      return Math.min(1.5, spendUpsideScore);
    }
    
    // No significant impression share lost to budget
    return 1.0; // Neutral score
  } catch (error) {
    Logger.log(`Error calculating spend upside score for campaign ${campaign.name}: ${error.message}`);
    return 1; // Return neutral score on error
  }
}

function calculateDayOfWeekScore(campaign, currentDayOfWeek, config) {
  try {
    // Get performance pattern with adaptive lookback
    const performancePattern = getDayOfWeekPerformancePattern(
      campaign, 
      config.LOOKBACK_PERIOD_DOW,
      config
    );
    
    // If we don't have confidence in the data, return neutral factor (1.0)
    if (!performancePattern.hasConfidence) {
      return {
        score: 1.0,
        confidence: false,
        daysAnalyzed: performancePattern.totalDaysAnalyzed,
        conversions: performancePattern.totalConversions
      };
    }
    
    // Return day-specific performance factor
    return {
      score: performancePattern.dayIndex[currentDayOfWeek],
      confidence: true,
      daysAnalyzed: performancePattern.totalDaysAnalyzed,
      conversions: performancePattern.totalConversions
    };
  } catch (error) {
    Logger.log(`Error calculating day of week score for campaign ${campaign.name}: ${error.message}`);
    
    // Return neutral score on error
    return {
      score: 1.0,
      confidence: false,
      daysAnalyzed: 0,
      conversions: 0,
      error: error.message
    };
  }
}

function buildStrategySpecificQuery(campaign, dateRange, strategyMetrics) {
  // Build appropriate AWQL query based on strategy metrics
  const volumeMetric = strategyMetrics.volumeMetric;
  const performanceMetric = strategyMetrics.performanceMetric;
  
  let metrics = ["Impressions", "Cost"];
  
  // Add the volume metric if not already included
  if (!metrics.includes(volumeMetric.charAt(0).toUpperCase() + volumeMetric.slice(1))) {
    metrics.push(volumeMetric.charAt(0).toUpperCase() + volumeMetric.slice(1));
  }
  
  // Add performance-specific metrics
  if (performanceMetric === "ctr") {
    metrics.push("Ctr");
  } else if (performanceMetric === "convRate") {
    metrics.push("ConversionRate");
  } else if (performanceMetric === "cpa") {
    metrics.push("CostPerConversion");
  } else if (performanceMetric === "roas") {
    metrics.push("ValuePerConversion");
  } else if (performanceMetric === "valuePerCost") {
    metrics.push("ConversionValue");
  } else if (performanceMetric === "impressionShare") {
    metrics.push("SearchImpressionShare", "SearchBudgetLostImpressionShare");
  } else if (performanceMetric === "roi") {
    metrics.push("ConversionValue");
  }
  
  // Build the query
  const query = `
    SELECT CampaignId, CampaignName, ${metrics.join(", ")}
    FROM CAMPAIGN_PERFORMANCE_REPORT
    WHERE CampaignId = ${campaign.getId()}
    AND Impressions > 0
    DURING ${dateRange.startDate},${dateRange.endDate}
  `;
  
  return query;
}

function validateConversionActions() {
  try {
    const allActions = new Set();
    
    // Collect all specified conversion actions
    if (CONFIG.CONVERSION_SETTINGS.MASTER_ALTERNATE_CONVERSION) {
      allActions.add(CONFIG.CONVERSION_SETTINGS.MASTER_ALTERNATE_CONVERSION);
    }
    
    Object.values(CONFIG.CONVERSION_SETTINGS.CAMPAIGN_MAPPINGS).forEach(action => {
      allActions.add(action);
    });
    
    // Validate each action exists
    const invalidActions = [];
    for (const action of allActions) {
      if (!conversionActionExists(action)) {
        invalidActions.push(action);
      }
    }
    
    if (invalidActions.length > 0) {
      const error = `Invalid conversion actions found: ${invalidActions.join(', ')}`;
      if (CONFIG.CONVERSION_SETTINGS.ON_INVALID_ACTION === 'ERROR') {
        throw new Error(error);
      } else {
        Logger.log(`WARNING: ${error}`);
      }
    }
    
    return true;
  } catch (error) {
    Logger.log(`Error validating conversion actions: ${error.message}`);
    return false;
  }
}

function generateSummaryReport(processedCampaigns, budgetData) {
  try {
    // ... existing summary stats ...
    
    // Add conversion source summary
    const conversionSources = {
      default: processedCampaigns.filter(c => c.conversionSource === 'default').length,
      alternate: processedCampaigns.filter(c => c.conversionSource === 'alternate').length
    };
    
    Logger.log("=== CONVERSION SOURCE SUMMARY ===");
    Logger.log(`Using Default Conversions: ${conversionSources.default}`);
    Logger.log(`Using Alternate Conversions: ${conversionSources.alternate}`);
    
    // Campaign-specific conversion details
    Logger.log("=== CAMPAIGN CONVERSION DETAILS ===");
    processedCampaigns.forEach(campaign => {
      Logger.log(`${campaign.name}:`);
      Logger.log(`  Conversion Source: ${campaign.conversionSource}`);
      if (campaign.conversionSource === 'alternate') {
        Logger.log(`  Conversion Action: ${campaign.conversionAction}`);
      }
    });
    
    // ... rest of reporting ...
  } catch (error) {
    Logger.log(`Error generating summary report: ${error.message}`);
  }
}
