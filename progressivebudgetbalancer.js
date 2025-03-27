// Update processCampaignData to properly handle conversions and shared budgets
function processCampaignData(campaign, dateRange, dowDateRange) {
  try {
    const stats = campaign.getStatsFor(dateRange.start, dateRange.end);
    const budget = campaign.getBudget();
    const campaignName = campaign.getName();
    
    // Get budget amount first
    const currentDailyBudget = budget.getAmount();
    
    // Detect shared budget status
    let isSharedBudget = false;
    let sharedBudgetId = null;
    
    try {
      const query = `
        SELECT campaign.id, campaign_budget.id, campaign_budget.type
        FROM campaign
        WHERE campaign.id = ${campaign.getId()}
      `;
      
      const report = AdsApp.report(query);
      const rows = report.rows();
      
      if (rows.hasNext()) {
        const row = rows.next();
        const budgetType = row['campaign_budget.type'];
        
        if (budgetType === 'SHARED') {
          isSharedBudget = true;
          sharedBudgetId = row['campaign_budget.id'];
        }
      }
    } catch (e) {
      // Fallback to name-based detection
      const budgetPattern = /BUDGET\s+\d+\s*\|/;
      isSharedBudget = budgetPattern.test(campaignName);
      
      if (isSharedBudget) {
        const match = campaignName.match(/BUDGET\s+(\d+)\s*\|/);
        if (match && match[1]) {
          sharedBudgetId = "group_" + match[1];
        }
      }
    }
    
    // Get metrics
    const cost = stats.getCost();
    const clicks = stats.getClicks();
    const impressions = stats.getImpressions();
    
    // Get conversion metrics
    let conversions, conversionValue;
    if (CONFIG.CONVERSION_SETTINGS.USE_SPECIFIC_CONVERSION_ACTION) {
      const conversionData = getSpecificConversionMetrics(
        campaign, 
        dateRange, 
        CONFIG.CONVERSION_SETTINGS.CONVERSION_ACTION_NAME
      );
      conversions = conversionData.conversions;
      conversionValue = conversionData.conversionValue;
    } else {
      conversions = stats.getConversions();
      conversionValue = conversions * CONFIG.CONVERSION_SETTINGS.ESTIMATED_CONVERSION_VALUE;
    }
    
    // Get recent conversions
    let recentConversions = 0;
    if (CONFIG.CONVERSION_SETTINGS.USE_SPECIFIC_CONVERSION_ACTION) {
      const endDate = new Date();
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - CONFIG.RECENT_PERFORMANCE_PERIOD);
      
      const recentDateRange = {
        start: Utilities.formatDate(startDate, 'UTC', 'yyyyMMdd'),
        end: Utilities.formatDate(endDate, 'UTC', 'yyyyMMdd')
      };
      
      const recentData = getSpecificConversionMetrics(
        campaign, 
        recentDateRange, 
        CONFIG.CONVERSION_SETTINGS.CONVERSION_ACTION_NAME
      );
      recentConversions = recentData.conversions;
    } else {
      const endDate = new Date();
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - CONFIG.RECENT_PERFORMANCE_PERIOD);
      
      const recentDateRange = {
        start: Utilities.formatDate(startDate, 'UTC', 'yyyyMMdd'),
        end: Utilities.formatDate(endDate, 'UTC', 'yyyyMMdd')
      };
      
      const recentStats = campaign.getStatsFor(recentDateRange.start, recentDateRange.end);
      recentConversions = recentStats.getConversions();
    }
    
    // Get performance metrics
    const impressionShare = getImpressionShare(campaign, dateRange);
    const budgetImpressionShareLost = getImpressionShareLostToBudget(campaign, dateRange);
    
    // Get day-of-week data
    let dayOfWeekData = null;
    let dayOfWeekAdjustment = {
      appliedMultiplier: 1.0,
      confidence: 0.0,
      rawMultiplier: 1.0,
      dayName: "Fallback"
    };
    
    if (CONFIG.DAY_OF_WEEK.ENABLED) {
      try {
        dayOfWeekData = getDayOfWeekData(campaign, dowDateRange);
        
        if (dayOfWeekData && dayOfWeekData.adjustment) {
          dayOfWeekAdjustment = dayOfWeekData.adjustment;
        }
      } catch (e) {
        Logger.log(`Error getting day-of-week data for ${campaignName}: ${e}`);
      }
    }
    
    // Create campaign data object
    return {
      campaign: campaign,
      name: campaignName,
      isSharedBudget: isSharedBudget,
      sharedBudgetId: sharedBudgetId,
      currentDailyBudget,
      cost,
      conversions,
      conversionValue,
      clicks,
      impressions,
      recentConversions,
      impressionShare,
      budgetImpressionShareLost,
      dayOfWeekData,
      dayOfWeekAdjustment
    };
  } catch (e) {
    Logger.log(`Error processing campaign ${campaign.getName()}: ${e}`);
    return null;
  }
}
