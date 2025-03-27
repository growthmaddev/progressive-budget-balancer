/**
 * Progressive Budget Balancer for Google Ads
 * 
 * This script progressively balances budgets over a three-month period to achieve
 * optimal alignment between budget allocation and conversion performance while
 * considering ROI and impression share metrics specific to campaign bid strategies.
 * 
 * This version compares recent (21-day) performance against the 90-day baseline
 * to make budget decisions more responsive to performance trends.
 */

// Configuration
const CONFIG = {
  // Lookback period in days (approximately 3 months)
  LOOKBACK_PERIOD: 90,
  
  // Recent performance lookback period in days
  RECENT_PERFORMANCE_PERIOD: 21,
  
  // PREVIEW MODE - Set to true to run without making any actual changes (for testing)
  PREVIEW_MODE: false,
  
  // Total monthly budget across all campaigns (in account currency)
  TOTAL_MONTHLY_BUDGET: 10000, // Will be divided by days in current month to get daily budget
  
  // Threshold for impression share lost due to budget (as a percentage)
  MAX_IMPRESSION_SHARE_LOST_TO_BUDGET: 10,
  
  // Minimum ROI threshold (as a multiplier, e.g., 3 = 300% ROI)
  MIN_ROI_THRESHOLD: 3,
  
  // Configuration for shared budget identification
  SHARED_BUDGET_IDENTIFIER: "BUDGET ", // Campaigns with this pattern in their name use shared budgets
  
  // Conversion tracking settings
  CONVERSION_SETTINGS: {
    // Set to true to use a specific conversion action instead of the default conversions
    USE_SPECIFIC_CONVERSION_ACTION: true,
    // The name of the specific conversion action to track (only applies if USE_SPECIFIC_CONVERSION_ACTION is true)
    CONVERSION_ACTION_NAME: "create portfolio",
    // Estimated value per conversion (only used if we can't get actual values)
    ESTIMATED_CONVERSION_VALUE: 10,
    // New settings
    MIN_CONVERSIONS_FOR_PATTERN: 5,  // Minimum conversions needed for pattern validation
    PATTERN_STRENGTH_THRESHOLD: 0.08, // Reduced from 0.1 for less restrictive validation
    CONFIDENCE_THRESHOLD: 0.25,      // Reduced from 0.3 for more adjustments
    VOLUME_BASED_THRESHOLDS: {
      LOW: { min: 0, threshold: 0.08 },
      MEDIUM: { min: 10, threshold: 0.1 },
      HIGH: { min: 50, threshold: 0.12 }
    }
  },
  
  // NEW: Day-of-week optimization settings
  DAY_OF_WEEK: {
    // Enable day-of-week optimization
    ENABLED: true,
    // Initial lookback period for day-of-week analysis (days)
    LOOKBACK_PERIOD: 90,
    // Minimum data points required for a day to be considered reliable
    MIN_DATA_POINTS: 8,
    // Maximum day-specific multiplier allowed (1.3 = up to 30% increase)
    MAX_DAY_MULTIPLIER: 1.3,
    // Minimum day-specific multiplier allowed (0.7 = up to 30% decrease)
    MIN_DAY_MULTIPLIER: 0.7,
    // Weight of day-of-week factor in overall budget decision (0-1)
    INFLUENCE_WEIGHT: 0.4, // Increased from 0.25 to 0.4 for stronger day-specific influence
    // Special events calendar to adjust for holidays, etc.
    SPECIAL_EVENTS: [
      // Each entry follows format: { date: "YYYY-MM-DD", name: "Event Name", multiplier: 1.0 }
      // Example: { date: "2023-11-24", name: "Black Friday", multiplier: 1.5 }
    ],
    // NEW: Adaptive lookback settings
    ADAPTIVE_LOOKBACK: {
      // Enable adaptive lookback period
      ENABLED: true,
      // Increment size in days for each extension (7 = weekly)
      INCREMENT_SIZE: 7,
      // Minimum confidence threshold to consider data reliable
      MIN_CONFIDENCE_THRESHOLD: 0.3,
      // Maximum number of extensions to try (set a reasonable limit)
      MAX_EXTENSIONS: 8, // Up to 8 additional weeks (8*7=56 days)
      // Maximum total lookback period in days (90 initial + 56 extended = 146 days max)
      MAX_TOTAL_LOOKBACK: 146,
      // Fallback confidence threshold to use when we can't find enough reliable data
      FALLBACK_CONFIDENCE_THRESHOLD: 0.15,
      // Enable progressive threshold relaxation (gradually lower threshold if needed)
      PROGRESSIVE_RELAXATION: true,
      // Relaxation steps - array of confidence thresholds to try in sequence
      RELAXATION_STEPS: [0.25, 0.2, 0.15, 0.1, 0.05],
      // NEW: Retrospective simulation settings
      RETROSPECTIVE_SIMULATION: {
        // Enable retrospective simulation for previous days
        ENABLED: true,
        // Number of previous days to simulate
        DAYS_TO_SIMULATE: 6,
        // Include weekend days in simulation
        INCLUDE_WEEKENDS: true
      }
    }
  },
  
  // Threshold values for different bid strategies
  STRATEGY_THRESHOLDS: {
    'TARGET_CPA': { threshold: 0, better_if_lower: true }, // Target CPA in account currency
    'TARGET_ROAS': { threshold: 400, better_if_lower: false }, // Target ROAS as percentage
    'MAXIMIZE_CONVERSIONS': { threshold: 0, better_if_lower: false }, // Any conversions is good
    'MAXIMIZE_CONVERSION_VALUE': { threshold: 0, better_if_lower: false }, // Any value is good
    'TARGET_IMPRESSION_SHARE': { threshold: 80, better_if_lower: false }, // Target impression share percentage
    'MAXIMIZE_CLICKS': { threshold: 0, better_if_lower: false }, // Any clicks is good
    'MANUAL_CPC': { threshold: 3, better_if_lower: false } // Default ROI threshold
  },
  
  // Percentage adjustment per iteration (smaller = more gradual changes)
  MAX_ADJUSTMENT_PERCENTAGE: 15,
  
  // Minimum percentage change required to apply a budget update
  MIN_ADJUSTMENT_PERCENTAGE: 0.1,
  
  // Frequency of script execution (for logging purposes)
  EXECUTION_FREQUENCY: 'DAILY', // Options: 'DAILY', 'WEEKLY', 'MONTHLY'
  
  // Performance optimization settings
  PERFORMANCE: {
    // Maximum number of campaigns to process (set to 0 for unlimited)
    MAX_CAMPAIGNS: 100,
    // Logging frequency (log every N campaigns)
    LOG_FREQUENCY: 10
  },
  
  // Budget pacing settings
  BUDGET_PACING: {
    // Minimum percentage of monthly budget that should be spent by end of month
    MIN_MONTHLY_SPEND_PERCENTAGE: 95,
    // Maximum percentage of monthly budget that can be spent by end of month
    MAX_MONTHLY_SPEND_PERCENTAGE: 105,
    // Minimum days of data required for pacing calculation
    MIN_DAYS_FOR_PACING: 3,
    // Whether to allow budget redistribution to high-performance days
    ALLOW_BUDGET_REDISTRIBUTION: true,
    // Maximum daily budget multiplier for high-performance days
    MAX_DAILY_BUDGET_MULTIPLIER: 1.5
  },
  
  // Add to CONFIG:
  SPECIAL_EVENTS: {
    // Example event configuration
    "Black Friday": {
      startDate: "2024-11-29",
      endDate: "2024-11-29",
      multiplier: 1.5,
      impactDays: 3,  // Days before and after to consider impact
      confidence: 1.0,
      description: "Black Friday sale"
    }
  },
  
  MEMORY_LIMITS: {
    MAX_CAMPAIGNS_PER_BATCH: 100,
    MAX_LOOKBACK_DAYS: 146,
    MAX_SPECIAL_EVENTS: 50,
    CACHE_EXPIRY_MS: 3600000, // 1 hour
    MAX_CACHE_SIZE: 1000
  }
};

// At the top of your script file, add this line:
var scriptTimezone; // Script-level variable to store account timezone

// Initialize caching system at the module level
let campaignDataCache = null;
let cacheTimestamps = null;

function initializeCache() {
  if (!campaignDataCache) {
    campaignDataCache = new Map();
    cacheTimestamps = new Map();
  }
}

function cleanupCache() {
  if (!cacheTimestamps) return;
  
  const now = Date.now();
  let cleanedCount = 0;
  
  for (const [key, timestamp] of cacheTimestamps.entries()) {
    if (now - timestamp > CONFIG.MEMORY_LIMITS.CACHE_EXPIRY_MS) {
      campaignDataCache.delete(key);
      cacheTimestamps.delete(key);
      cleanedCount++;
    }
  }
  
  if (cleanedCount > 0) {
    Logger.log(`Cleaned ${cleanedCount} expired cache entries`);
  }
}

function getCachedCampaignData(campaign, dateRange) {
  try {
    initializeCache();
    
    const cacheKey = `${campaign.getId()}-${dateRange.start}-${dateRange.end}`;
    const now = Date.now();
    const cachedTimestamp = cacheTimestamps.get(cacheKey);
    
    // Check if cache exists and is not expired
    if (campaignDataCache.has(cacheKey) && 
        cachedTimestamp && 
        (now - cachedTimestamp) < CONFIG.MEMORY_LIMITS.CACHE_EXPIRY_MS) {
      return campaignDataCache.get(cacheKey);
    }
    
    // If cache is full, remove oldest entries
    if (campaignDataCache.size >= CONFIG.MEMORY_LIMITS.MAX_CACHE_SIZE) {
      cleanupCache();
    }
    
    // Get fresh data
    const freshData = collectCampaignData(dateRange);
    campaignDataCache.set(cacheKey, freshData);
    cacheTimestamps.set(cacheKey, now);
    
    return freshData;
  } catch (e) {
    Logger.log(`Error in getCachedCampaignData for ${campaign.getName()}: ${e}`);
    return collectCampaignData(dateRange);
  }
}

/**
 * Calculate the maximum daily budget based on the current month
 * @return {number} The calculated maximum daily budget
 */
function calculateMaxDailyBudget() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth(); // 0-11
  
  // Get the number of days in the current month
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  
  // Calculate daily budget
  const dailyBudget = CONFIG.TOTAL_MONTHLY_BUDGET / daysInMonth;
  
  Logger.log(`\n=== Monthly Budget Calculation ===`);
  Logger.log(`Total monthly budget: ${CONFIG.TOTAL_MONTHLY_BUDGET}`);
  Logger.log(`Days in current month: ${daysInMonth}`);
  Logger.log(`Calculated daily budget: ${dailyBudget.toFixed(2)}`);
  Logger.log(`=====================================\n`);
  
  return dailyBudget;
}

/**
 * Day-of-Week Performance Analysis
 * 
 * These functions analyze campaign performance by day of week to create
 * day-specific performance indices for budget optimization.
 */

// Get performance data by day of week for a campaign
function getDayOfWeekPerformance(campaign, dateRange) {
  try {
    // Create simplified object to store only conversion data by day
    const dayPerformance = {
      0: { name: 'Sunday', conversions: 0, days: 0 },
      1: { name: 'Monday', conversions: 0, days: 0 },
      2: { name: 'Tuesday', conversions: 0, days: 0 },
      3: { name: 'Wednesday', conversions: 0, days: 0 },
      4: { name: 'Thursday', conversions: 0, days: 0 },
      5: { name: 'Friday', conversions: 0, days: 0 },
      6: { name: 'Saturday', conversions: 0, days: 0 }
    };
    
    // Get conversion data using the proper method for specific conversion actions
    let query;
    if (CONFIG.CONVERSION_SETTINGS.USE_SPECIFIC_CONVERSION_ACTION) {
      // Use the specialized query for specific conversion actions
      query = `
        SELECT 
          segments.date,
          segments.conversion_action_name,
          metrics.all_conversions
        FROM campaign 
        WHERE 
          campaign.id = ${campaign.getId()} 
          AND segments.date BETWEEN "${dateRange.start}" AND "${dateRange.end}"
      `;
    } else {
      // For default conversions, use simple query with regular conversions
      query = `
        SELECT 
          segments.date,
          metrics.conversions
        FROM campaign 
        WHERE 
          campaign.id = ${campaign.getId()} 
          AND segments.date BETWEEN "${dateRange.start}" AND "${dateRange.end}"
      `;
    }

    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Process each day's data
    while (rows.hasNext()) {
      const row = rows.next();
      
      // Get date and convert to day of week
      const dateString = row['segments.date'];
      const date = new Date(
        parseInt(dateString.substring(0, 4)),
        parseInt(dateString.substring(4, 6)) - 1,
        parseInt(dateString.substring(6, 8))
      );
      const dayOfWeek = date.getDay();
      
      // Only count conversions for the specific action if configured
      let conversions = 0;
      if (CONFIG.CONVERSION_SETTINGS.USE_SPECIFIC_CONVERSION_ACTION) {
        const actionName = row['segments.conversion_action_name'] || '';
        if (actionName === CONFIG.CONVERSION_SETTINGS.CONVERSION_ACTION_NAME) {
          conversions = parseFloat(row['metrics.all_conversions']) || 0;
        }
      } else {
        conversions = parseFloat(row['metrics.conversions']) || 0;
      }
      
      // Add to the appropriate day
      dayPerformance[dayOfWeek].conversions += conversions;
      dayPerformance[dayOfWeek].days += 1;
    }
    
    // Calculate average daily conversions for each day
    for (let i = 0; i < 7; i++) {
      const day = dayPerformance[i];
      if (day.days > 0) {
        day.avgDailyConversions = day.conversions / day.days;
      }
    }
    
    return dayPerformance;
  } catch (e) {
    Logger.log(`Error in getDayOfWeekPerformance for campaign ${campaign.getName()}: ${e}`);
    return null;
  }
}

// Calculate day-of-week performance indices for a campaign
function calculateDayOfWeekAdjustment(dayIndices) {
  try {
    // Safety check - if dayIndices is null or undefined, return neutral
    if (!dayIndices) {
      return {
        dayName: "Unknown",
        rawMultiplier: 1.0,
        appliedMultiplier: 1.0,
        confidence: 0,
        isSpecialEvent: false,
        patternValidation: { isSignificant: false, reason: "No data" }
      };
    }
    
    // Get today's date in the account's timezone
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    const today = new Date();
    const todayInAccountTimezone = Utilities.formatDate(today, accountTimezone, "yyyy-MM-dd");
    
    // Get day of week in account timezone (0-6, Sunday-Saturday)
    const dayOfWeek = parseInt(Utilities.formatDate(today, accountTimezone, "u")) % 7;
    
    // Check if a special event is configured for today
    const todayString = Utilities.formatDate(today, 'UTC', 'yyyy-MM-dd');
    let specialEvent = null;
    
    // Check if CONFIG.DAY_OF_WEEK.SPECIAL_EVENTS exists before iterating
    if (CONFIG.DAY_OF_WEEK.SPECIAL_EVENTS && Array.isArray(CONFIG.DAY_OF_WEEK.SPECIAL_EVENTS)) {
      for (const event of CONFIG.DAY_OF_WEEK.SPECIAL_EVENTS) {
        if (event.date === todayString) {
          specialEvent = event;
          break;
        }
      }
    }
    
    // If it's a special event, use that multiplier instead
    if (specialEvent) {
      Logger.log(`Today is a special event: ${specialEvent.name} with multiplier ${specialEvent.multiplier}`);
      return {
        dayName: `${dayIndices[dayOfWeek] ? dayIndices[dayOfWeek].name : "Unknown"} (${specialEvent.name})`,
        rawMultiplier: specialEvent.multiplier,
        appliedMultiplier: specialEvent.multiplier,
        confidence: 1.0,
        isSpecialEvent: true,
        specialEventName: specialEvent.name,
        patternValidation: { isSignificant: true, reason: "Special event" }
      };
    }
    
    // Get the day's performance data with safe access
    const dayIndex = dayIndices[dayOfWeek];
    
    // If we don't have data for today's day of week or not enough data points, return neutral
    if (!dayIndex || !dayIndex.sampleSize || dayIndex.sampleSize < CONFIG.DAY_OF_WEEK.MIN_DATA_POINTS) {
      return {
        dayName: dayIndex ? dayIndex.name : "Unknown",
        rawMultiplier: 1.0,
        appliedMultiplier: 1.0,
        confidence: 0,
        isSpecialEvent: false,
        patternValidation: { isSignificant: false, reason: "Insufficient data" }
      };
    }
    
    // Validate the day-of-week pattern
    // Ensure the validateDayOfWeekPattern function exists and handles null/undefined values
    const patternValidation = validateDayOfWeekPattern(dayIndices, dayIndex.conversionIndex || 0);
    
    // If pattern is not significant, return neutral multiplier
    if (!patternValidation.isSignificant) {
      Logger.log(`Day-of-week pattern not significant: ${patternValidation.reason}`);
      return {
        dayName: dayIndex.name,
        rawMultiplier: 1.0,
        appliedMultiplier: 1.0,
        confidence: 0,
        isSpecialEvent: false,
        patternValidation
      };
    }
    
    // Calculate multiplier based on conversion performance
    let rawMultiplier = dayIndex.conversionIndex || 1.0;
    
    // Adjust multiplier based on confidence and pattern strength
    const confidenceAdjustedMultiplier = 1.0 + 
      ((rawMultiplier - 1.0) * dayIndex.confidence * patternValidation.patternStrength);
    
    // Constrain multiplier to configured limits
    const constrainedMultiplier = Math.max(
      CONFIG.DAY_OF_WEEK.MIN_DAY_MULTIPLIER,
      Math.min(CONFIG.DAY_OF_WEEK.MAX_DAY_MULTIPLIER, confidenceAdjustedMultiplier)
    );
    
    return {
      dayName: dayIndex.name,
      rawMultiplier: rawMultiplier,
      appliedMultiplier: constrainedMultiplier,
      confidence: dayIndex.confidence,
      isSpecialEvent: false,
      patternValidation
    };
  } catch (e) {
    Logger.log(`Error calculating day-of-week adjustment: ${e}`);
    return {
      dayName: "Error",
      rawMultiplier: 1.0,
      appliedMultiplier: 1.0,
      confidence: 0,
      isSpecialEvent: false,
      patternValidation: { isSignificant: false, reason: "Error in calculation" }
    };
  }
}

/**
 * Validates if a day-of-week pattern is statistically significant
 * @param {Object} dayIndices - The day indices object containing performance data
 * @return {Object} - Validation results including significance and pattern strength
 */
function validateDayOfWeekPattern(dayIndices, conversionVolume) {
  try {
    // Safety check - if dayIndices is null or undefined, return not significant
    if (!dayIndices) {
      return {
        isSignificant: false,
        patternStrength: 0,
        reason: "No day indices data available"
      };
    }
    
    // Ensure conversion volume is a number
    conversionVolume = parseFloat(conversionVolume) || 0;
    
    // Get volume-based threshold
    let threshold = CONFIG.CONVERSION_SETTINGS.VOLUME_BASED_THRESHOLDS.LOW.threshold;
    if (conversionVolume >= CONFIG.CONVERSION_SETTINGS.VOLUME_BASED_THRESHOLDS.HIGH.min) {
      threshold = CONFIG.CONVERSION_SETTINGS.VOLUME_BASED_THRESHOLDS.HIGH.threshold;
    } else if (conversionVolume >= CONFIG.CONVERSION_SETTINGS.VOLUME_BASED_THRESHOLDS.MEDIUM.min) {
      threshold = CONFIG.CONVERSION_SETTINGS.VOLUME_BASED_THRESHOLDS.MEDIUM.threshold;
    }
    
    // Calculate pattern significance - safely handle the dayIndices object
    const indices = Object.values(dayIndices)
      .filter(day => day && typeof day === 'object' && 
              day.sampleSize >= CONFIG.CONVERSION_SETTINGS.MIN_CONVERSIONS_FOR_PATTERN &&
              typeof day.conversionIndex === 'number')
      .map(day => day.conversionIndex);
    
    if (indices.length < 3) {
      return {
        isSignificant: false,
        patternStrength: 0,
        reason: "Insufficient days with data"
      };
    }
    
    const mean = indices.reduce((sum, val) => sum + val, 0) / indices.length;
    const variance = indices.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / indices.length;
    const stdDev = Math.sqrt(variance);
    const patternStrength = mean > 0 ? stdDev / mean : 0;
    
    const isSignificant = patternStrength >= threshold && indices.length >= 3;
    
    return {
      isSignificant,
      patternStrength,
      stdDev,
      mean,
      threshold,
      reason: isSignificant ? 
        `Pattern is significant (strength: ${patternStrength.toFixed(2)})` :
        `Pattern is not significant (strength: ${patternStrength.toFixed(2)})`
    };
  } catch (e) {
    Logger.log(`Error validating day-of-week pattern: ${e}`);
    return {
      isSignificant: false,
      patternStrength: 0,
      reason: "Error in validation: " + e.message
    };
  }
}

// Modify calculateDayOfWeekAdjustment to use pattern validation
function calculateDayOfWeekAdjustment(dayIndices) {
  try {
    // Get today's date in the account's timezone
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    const today = new Date();
    const todayInAccountTimezone = Utilities.formatDate(today, accountTimezone, "yyyy-MM-dd");
    
    // Get day of week in account timezone (0-6, Sunday-Saturday)
    const dayOfWeek = parseInt(Utilities.formatDate(today, accountTimezone, "u")) % 7;
    
    // Check if a special event is configured for today
    const todayString = Utilities.formatDate(today, 'UTC', 'yyyy-MM-dd');
    let specialEvent = null;
    
    for (const event of CONFIG.DAY_OF_WEEK.SPECIAL_EVENTS) {
      if (event.date === todayString) {
        specialEvent = event;
        break;
      }
    }
    
    // If it's a special event, use that multiplier instead
    if (specialEvent) {
      Logger.log(`Today is a special event: ${specialEvent.name} with multiplier ${specialEvent.multiplier}`);
      return {
        dayName: `${dayIndices[dayOfWeek].name} (${specialEvent.name})`,
        rawMultiplier: specialEvent.multiplier,
        appliedMultiplier: specialEvent.multiplier,
        confidence: 1.0,
        isSpecialEvent: true,
        specialEventName: specialEvent.name,
        patternValidation: { isSignificant: true, reason: "Special event" }
      };
    }
    
    // Get the day's performance data
    const dayIndex = dayIndices[dayOfWeek];
    
    // If we don't have enough data points, return neutral multiplier
    if (!dayIndex || dayIndex.sampleSize < CONFIG.DAY_OF_WEEK.MIN_DATA_POINTS) {
      return {
        dayName: dayIndex ? dayIndex.name : "Unknown",
        rawMultiplier: 1.0,
        appliedMultiplier: 1.0,
        confidence: 0,
        isSpecialEvent: false,
        patternValidation: { isSignificant: false, reason: "Insufficient data" }
      };
    }
    
    // Validate the day-of-week pattern
    const patternValidation = validateDayOfWeekPattern(dayIndices, dayIndex.conversionIndex);
    
    // If pattern is not significant, return neutral multiplier
    if (!patternValidation.isSignificant) {
      Logger.log(`Day-of-week pattern not significant: ${patternValidation.reason}`);
      return {
        dayName: dayIndex.name,
        rawMultiplier: 1.0,
        appliedMultiplier: 1.0,
        confidence: 0,
        isSpecialEvent: false,
        patternValidation
      };
    }
    
    // Calculate multiplier based on conversion performance
    let rawMultiplier = dayIndex.conversionIndex;
    
    // Adjust multiplier based on confidence and pattern strength
    const confidenceAdjustedMultiplier = 1.0 + 
      ((rawMultiplier - 1.0) * dayIndex.confidence * patternValidation.patternStrength);
    
    // Constrain multiplier to configured limits
    const constrainedMultiplier = Math.max(
      CONFIG.DAY_OF_WEEK.MIN_DAY_MULTIPLIER,
      Math.min(CONFIG.DAY_OF_WEEK.MAX_DAY_MULTIPLIER, confidenceAdjustedMultiplier)
    );
    
    return {
      dayName: dayIndex.name,
      rawMultiplier: rawMultiplier,
      appliedMultiplier: constrainedMultiplier,
      confidence: dayIndex.confidence,
      isSpecialEvent: false,
      patternValidation
    };
  } catch (e) {
    Logger.log(`Error calculating day-of-week adjustment: ${e}`);
    return {
      dayName: "Error",
      rawMultiplier: 1.0,
      appliedMultiplier: 1.0,
      confidence: 0,
      isSpecialEvent: false,
      patternValidation: { isSignificant: false, reason: "Error in calculation" }
    };
  }
}

/**
 * Gets day-of-week performance data with adaptive lookback period
 * This function will extend the lookback period as needed until it finds enough reliable data
 * for the current day of the week or reaches the maximum allowed lookback period.
 */
function getAdaptiveDayOfWeekData(campaign, purpose = 'main') {
  try {
    const campaignName = campaign.getName();
    
    // Get today's date in the account's timezone
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    const today = new Date();
    const todayInAccountTimezone = Utilities.formatDate(today, accountTimezone, "yyyy-MM-dd");
    
    // Get day of week in account timezone (0-6, Sunday-Saturday)
    const dayOfWeek = parseInt(Utilities.formatDate(today, accountTimezone, "u")) % 7;
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const todayName = dayNames[dayOfWeek];
    
    // Initialize variables for tracking lookback attempts
    let currentLookback = CONFIG.DAY_OF_WEEK.LOOKBACK_PERIOD;
    let extensionsUsed = 0;
    let lookbackAttempts = [];
    let usingRelaxedThreshold = false;
    let fallbackApplied = false;
    
    // Try progressive relaxation of confidence threshold if enabled
    const confidenceThresholds = CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.PROGRESSIVE_RELAXATION ? 
      [CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MIN_CONFIDENCE_THRESHOLD, ...CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.RELAXATION_STEPS] :
      [CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MIN_CONFIDENCE_THRESHOLD];
    
    // Log the start of analysis for this campaign with purpose
    Logger.log(`\n=== Day-of-Week Analysis for ${campaignName} (${purpose}) ===`);
    Logger.log(`Analyzing performance for ${todayName} (${todayInAccountTimezone})`);
    
    for (const threshold of confidenceThresholds) {
      while (currentLookback <= CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MAX_TOTAL_LOOKBACK && 
             extensionsUsed < CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MAX_EXTENSIONS) {
        
        // Calculate date range for current lookback period
        const endDate = new Date();
        const startDate = new Date();
        startDate.setDate(startDate.getDate() - currentLookback);
        
        const dateRange = {
          start: Utilities.formatDate(startDate, 'UTC', 'yyyyMMdd'),
          end: Utilities.formatDate(endDate, 'UTC', 'yyyyMMdd')
        };
        
        // Get performance data for this lookback period
        const performance = getDayOfWeekPerformance(campaign, dateRange);
        
        // Safely calculate indices if performance data exists
        let indices = null;
        if (performance) {
          indices = calculateDayOfWeekIndices(performance);
        } else {
          // Skip to next iteration if no performance data
          lookbackAttempts.push({
            lookback: currentLookback,
            threshold: threshold,
            success: false,
            reason: "No performance data"
          });
          
          currentLookback += CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.INCREMENT_SIZE;
          extensionsUsed++;
          continue;
        }
        
        // Safely check if we have indices and day-specific data
        if (!indices || !indices.dayIndices) {
          lookbackAttempts.push({
            lookback: currentLookback,
            threshold: threshold,
            success: false,
            reason: "No indices data"
          });
          
          currentLookback += CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.INCREMENT_SIZE;
          extensionsUsed++;
          continue;
        }
        
        // Safely get today's data
        const todayData = indices.dayIndices[dayOfWeek];
        
        // Check if we have enough data points for today
        if (todayData && todayData.sampleSize >= CONFIG.DAY_OF_WEEK.MIN_DATA_POINTS) {
          // Calculate confidence based on sample size
          const confidence = Math.min(1.0, todayData.sampleSize / 30);
          
          if (confidence >= threshold) {
            // We found reliable data
            Logger.log(`\nAnalysis Results:`);
            Logger.log(`  - Sample size: ${todayData.sampleSize} data points`);
            Logger.log(`  - Confidence: ${confidence.toFixed(2)} (threshold: ${threshold})`);
            Logger.log(`  - Lookback period: ${currentLookback} days`);
            Logger.log(`  - Raw multiplier: ${todayData.conversionIndex.toFixed(2)}`);
            
            // Calculate day multiplier
            const rawMultiplier = todayData.conversionIndex;
            const dayMultiplier = Math.min(
              CONFIG.DAY_OF_WEEK.MAX_DAY_MULTIPLIER,
              Math.max(CONFIG.DAY_OF_WEEK.MIN_DAY_MULTIPLIER, rawMultiplier)
            );
            
            Logger.log(`  - Applied multiplier: ${dayMultiplier.toFixed(2)}`);
            
            // Return the successful data
            return {
              performance: performance,
              indices: indices,
              adjustment: {
                appliedMultiplier: dayMultiplier,
                confidence: confidence,
                rawMultiplier: rawMultiplier,
                dayName: todayName
              },
              lookbackUsed: currentLookback,
              extensionsUsed: extensionsUsed,
              lookbackAttempts: lookbackAttempts,
              usingRelaxedThreshold: threshold !== CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MIN_CONFIDENCE_THRESHOLD,
              confidenceThreshold: threshold,
              fallbackApplied: false,
              analysisPurpose: purpose
            };
          }
        }
        
        // Record this attempt
        lookbackAttempts.push({
          lookback: currentLookback,
          threshold: threshold,
          success: false,
          reason: todayData ? 
            `Insufficient confidence: ${todayData.sampleSize} samples` : 
            "No data for today"
        });
        
        // Extend lookback period
        currentLookback += CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.INCREMENT_SIZE;
        extensionsUsed++;
      }
    }
    
    // If we get here, we couldn't find reliable data with any threshold
    Logger.log(`\nUnable to find reliable data for ${todayName}`);
    Logger.log(`===========================================\n`);
    
    // Return default values with error information
    return {
      performance: null,
      indices: null,
      adjustment: {
        appliedMultiplier: 1.0,
        confidence: 0.0,
        rawMultiplier: 1.0,
        dayName: todayName
      },
      lookbackUsed: currentLookback,
      extensionsUsed: extensionsUsed,
      lookbackAttempts: lookbackAttempts,
      usingRelaxedThreshold: false,
      confidenceThreshold: CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MIN_CONFIDENCE_THRESHOLD,
      fallbackApplied: true,
      analysisPurpose: purpose,
      error: "Unable to find reliable data"
    };
    
  } catch (e) {
    Logger.log(`Error in day-of-week analysis for campaign ${campaign.getName()}: ${e}`);
    // Return a safe default object even when errors occur
    return {
      performance: null,
      indices: null,
      adjustment: {
        appliedMultiplier: 1.0,
        confidence: 0.0,
        rawMultiplier: 1.0,
        dayName: "Error"
      },
      error: e.toString(),
      fallbackApplied: true
    };
  }
}

function main() {
  try {
    // Initialize script
    initializeScript();
    
    // Get date ranges
    const dateRange = getDateRange();
    const dowDateRange = getDayOfWeekDateRange();
    
    // Log configuration
    logConfiguration(dateRange);
    
    // Test conversion action availability
    testConversionActionAvailability();
    
    // Collect campaign data with error handling
    let campaignData = null;
    try {
      campaignData = collectCampaignData(dateRange, dowDateRange);
      if (!campaignData || !campaignData.campaigns) {
        throw new Error("Failed to collect campaign data");
      }
    } catch (e) {
      Logger.log(`Error collecting campaign data: ${e}`);
      return;
    }
    
    // Calculate trend factors for all campaigns - CRITICAL STEP
    try {
      calculateTrendFactorsForAll(campaignData);
    } catch (e) {
      Logger.log(`Error calculating trend factors: ${e}`);
    }
    
    // Log current budget status
    logCurrentBudgetStatus(campaignData);

    // Calculate budget pacing information
    let pacingInfo = null;
    try {
      pacingInfo = calculateBudgetPacing();
    } catch (e) {
      Logger.log(`Error calculating budget pacing: ${e}`);
    }
    
    // Process campaign budgets
    try {
      processCampaignBudgets(campaignData, dateRange, pacingInfo);
    } catch (e) {
      Logger.log(`Error processing campaign budgets: ${e}`);
      return;
    }
    
    // Log final status
    Logger.log("\nScript execution completed successfully.");
    
  } catch (e) {
    Logger.log(`Fatal error in main: ${e}`);
  }
}

function formatDateRange(startDate, endDate) {
  return {
    start: Utilities.formatDate(startDate, 'UTC', 'yyyyMMdd'),
    end: Utilities.formatDate(endDate, 'UTC', 'yyyyMMdd')
  };
}

function collectCampaignData(dateRange, dowDateRange) {
  try {
    // Use cached data if available
    const cacheKey = `${dateRange.start}-${dateRange.end}`;
    if (campaignDataCache && campaignDataCache.has(cacheKey)) {
      return campaignDataCache.get(cacheKey);
    }
    
    // Collect fresh data
    const data = {
      campaigns: [],
      totalBudget: 0,
      totalSpend: 0,
      totalConversions: 0,
      dayOfWeekData: {}
    };
    
    // Process campaigns in batches to avoid memory issues
    const campaigns = AdsApp.campaigns()
      .withCondition("Status = ENABLED") // Changed from 'status' to 'Status'
      .get();
    
    let processedCount = 0;
    const batchSize = CONFIG.MEMORY_LIMITS.MAX_CAMPAIGNS_PER_BATCH;
    
    while (campaigns.hasNext()) {
      const campaign = campaigns.next();
      if (processedCount >= batchSize) break;
      
      try {
        const campaignData = processCampaignData(campaign, dateRange, dowDateRange);
        if (campaignData) {
          data.campaigns.push(campaignData);
          data.totalBudget += campaignData.currentDailyBudget;
          data.totalSpend += campaignData.cost;
          data.totalConversions += campaignData.conversions;
        }
      } catch (e) {
        Logger.log(`Error processing campaign ${campaign.getName()}: ${e}`);
        continue;
      }
      
      processedCount++;
    }
    
    // Cache the results
    if (campaignDataCache) {
      campaignDataCache.set(cacheKey, data);
      cacheTimestamps.set(cacheKey, Date.now());
    }
    
    return data;
  } catch (e) {
    Logger.log(`Error in collectCampaignData: ${e}`);
    return {
      campaigns: [],
      totalBudget: 0,
      totalSpend: 0,
      totalConversions: 0,
      dayOfWeekData: {}
    };
  }
}

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

function logDayOfWeekPerformanceSummary(campaignData) {
  // Get account timezone
  const accountTimezone = AdsApp.currentAccount().getTimeZone();
  
  // Format today's date in account timezone
  const today = new Date();
  const dayOfWeek = parseInt(Utilities.formatDate(today, accountTimezone, "u")) % 7;
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  Logger.log("\n===== DAY-OF-WEEK PERFORMANCE SUMMARY =====");
  Logger.log(`Day-of-week performance analysis for today (${dayNames[dayOfWeek]}):`);
  
  // Count of campaigns with reliable data
  const campaignsWithReliableData = campaignData.campaigns.filter(
    c => c.dayOfWeekAdjustment && c.dayOfWeekAdjustment.confidence >= 0.3
  ).length;
  
  Logger.log(`Campaigns with reliable day-of-week data: ${campaignsWithReliableData} of ${campaignData.campaigns.length}`);
  
  // Average confidence
  const avgConfidence = campaignData.campaigns.reduce(
    (sum, c) => sum + (c.dayOfWeekAdjustment ? c.dayOfWeekAdjustment.confidence : 0), 
    0
  ) / campaignData.campaigns.length;
  
  Logger.log(`Average day-of-week confidence: ${(avgConfidence * 100).toFixed(1)}%`);
  
  // Distribution of multipliers
  const multipliers = {
    increase: campaignData.campaigns.filter(c => 
      c.dayOfWeekAdjustment && c.dayOfWeekAdjustment.appliedMultiplier > 1.05
    ).length,
    decrease: campaignData.campaigns.filter(c => 
      c.dayOfWeekAdjustment && c.dayOfWeekAdjustment.appliedMultiplier < 0.95
    ).length,
    neutral: campaignData.campaigns.filter(c => 
      !c.dayOfWeekAdjustment || 
      (c.dayOfWeekAdjustment.appliedMultiplier >= 0.95 && c.dayOfWeekAdjustment.appliedMultiplier <= 1.05)
    ).length
  };
  
  Logger.log(`Multiplier distribution: ${multipliers.increase} increases, ${multipliers.decrease} decreases, ${multipliers.neutral} neutral`);
  Logger.log("===========================================\n");
}

function safeCalculateAdjustment(campaign, factors) {
  try {
    if (!campaign) {
      return {
        adjustmentFactor: 1.0,
        confidence: 0.0,
        reason: "Invalid input data"
      };
    }

    // Get adjustment factors directly from campaign object
    const trendFactor = campaign.trendFactor || 1.0;
    
    // Safely get day-of-week adjustment with fallback
    const dayOfWeekAdjustment = campaign.dayOfWeekAdjustment || { 
      appliedMultiplier: 1.0, 
      confidence: 0.0 
    };

    // Calculate combined adjustment factor
    let adjustmentFactor = 1.0;
    let confidence = 0.0;
    let reasons = [];

    // Apply trend factor
    if (trendFactor && !isNaN(trendFactor)) {
      adjustmentFactor *= trendFactor;
      reasons.push(`Trend factor: ${trendFactor.toFixed(2)}`);
      confidence += 0.3;
    }

    // Apply day-of-week adjustment only if confidence is sufficient
    if (dayOfWeekAdjustment && dayOfWeekAdjustment.appliedMultiplier && 
        dayOfWeekAdjustment.confidence >= 0.3) {
      adjustmentFactor *= dayOfWeekAdjustment.appliedMultiplier;
      reasons.push(`Day-of-week multiplier: ${dayOfWeekAdjustment.appliedMultiplier.toFixed(2)}`);
      confidence = Math.max(confidence, dayOfWeekAdjustment.confidence || 0.0);
    } else {
      reasons.push("No reliable day-of-week data available");
    }

    // Apply impression share lost adjustment when applicable
    if (campaign.budgetImpressionShareLost > 10) {
      const lossMultiplier = 1 + Math.min(0.15, campaign.budgetImpressionShareLost / 100);
      adjustmentFactor *= lossMultiplier;
      reasons.push(`Impression share lost: ${campaign.budgetImpressionShareLost.toFixed(1)}%`);
      confidence += 0.2;
    }

    // Ensure adjustment factor stays within bounds
    adjustmentFactor = Math.max(0.7, Math.min(1.3, adjustmentFactor));
    confidence = Math.min(1.0, confidence);

    return {
      adjustmentFactor: adjustmentFactor,
      confidence: confidence,
      reason: reasons.join(", ")
    };
  } catch (e) {
    Logger.log(`Error calculating adjustment: ${e}`);
    return {
      adjustmentFactor: 1.0,
      confidence: 0.0,
      reason: "Error in calculation: " + e.message
    };
  }
}

function getDayOfWeekData(campaign, dowDateRange) {
  try {
    if (CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.ENABLED) {
      const adaptiveData = getAdaptiveDayOfWeekData(campaign, 'main');
      
      // Check if adaptiveData is null (error occurred)
      if (!adaptiveData) {
        Logger.log(`Warning: No adaptive day data available for ${campaign.getName()}`);
        return {
          performance: null,
          indices: null,
          adjustment: {
            appliedMultiplier: 1.0,
            confidence: 0.0,
            rawMultiplier: 1.0,
            dayName: "Unknown"
          }
        };
      }
      
      return {
        performance: adaptiveData.performance,
        indices: adaptiveData.indices,
        adjustment: adaptiveData.adjustment || {
          appliedMultiplier: 1.0,
          confidence: 0.0,
          rawMultiplier: 1.0,
          dayName: "Fallback"
        },
        lookbackInfo: {
          lookbackUsed: adaptiveData.lookbackUsed || 0,
          reliableData: !adaptiveData.fallbackApplied
        }
      };
    } else {
      // Standard non-adaptive approach
      const performance = getDayOfWeekPerformance(campaign, dowDateRange);
      
      // Handle case where performance data is null
      if (!performance) {
        Logger.log(`Warning: No day-of-week performance data for ${campaign.getName()}`);
        return {
          performance: null,
          indices: null,
          adjustment: {
            appliedMultiplier: 1.0,
            confidence: 0.0,
            rawMultiplier: 1.0,
            dayName: "Unknown"
          }
        };
      }
      
      const indices = calculateDayOfWeekIndices(performance);
      
      // Handle case where indices calculation fails
      if (!indices || !indices.dayIndices) {
        Logger.log(`Warning: Could not calculate indices for ${campaign.getName()}`);
        return {
          performance: performance,
          indices: null,
          adjustment: {
            appliedMultiplier: 1.0,
            confidence: 0.0,
            rawMultiplier: 1.0,
            dayName: "Unknown"
          }
        };
      }
      
      const adjustment = calculateDayOfWeekAdjustment(indices.dayIndices);
      
      return { 
        performance, 
        indices, 
        adjustment: adjustment || {
          appliedMultiplier: 1.0,
          confidence: 0.0,
          rawMultiplier: 1.0,
          dayName: "Fallback"
        }
      };
    }
  } catch (e) {
    Logger.log(`Error getting day-of-week data for ${campaign.getName()}: ${e}`);
    // Return safe defaults
    return {
      performance: null,
      indices: null,
      adjustment: {
        appliedMultiplier: 1.0,
        confidence: 0.0,
        rawMultiplier: 1.0,
        dayName: "Error"
      },
      error: e.toString()
    };
  }
}

function getImpressionShareLostToBudget(campaign, dateRange) {
  try {
    // Create a query to get impression share metrics
    const query = 'SELECT campaign.id, metrics.search_budget_lost_impression_share ' +
      'FROM campaign ' +
      'WHERE campaign.id = ' + campaign.getId() + ' ' +
      'AND segments.date BETWEEN "' + dateRange.start + '" AND "' + dateRange.end + '"';
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Get the impression share lost due to budget
    if (rows.hasNext()) {
      const row = rows.next();
      return parseFloat(row['metrics.search_budget_lost_impression_share']) * 100 || 0;
    }
  } catch (e) {
    Logger.log("Error getting impression share lost for campaign " + campaign.getName() + ": " + e);
  }
  
  return 0;
}

function getImpressionShare(campaign, dateRange) {
  try {
    // Create a query to get impression share metrics
    const query = 'SELECT campaign.id, metrics.search_impression_share ' +
      'FROM campaign ' +
      'WHERE campaign.id = ' + campaign.getId() + ' ' +
      'AND segments.date BETWEEN "' + dateRange.start + '" AND "' + dateRange.end + '"';
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    // Get the impression share
    if (rows.hasNext()) {
      const row = rows.next();
      return parseFloat(row['metrics.search_impression_share']) * 100 || 0;
    }
  } catch (e) {
    Logger.log("Error getting impression share for campaign " + campaign.getName() + ": " + e);
  }
  
  return 0;
}

function calculateTrendFactor(campaign) {
  try {
    // Get recent conversions (21 days)
    const recentConvs = campaign.recentConversions;
    
    // Scale to equivalent 90-day rate for comparison
    const recentRate = recentConvs * (CONFIG.LOOKBACK_PERIOD / CONFIG.RECENT_PERFORMANCE_PERIOD);
    
    // Calculate conversion volume trend (>1 means improving, <1 means declining)
    let conversionTrend = 1.0; // Default to neutral
    
    if (campaign.conversions > 0) {
      conversionTrend = Math.min(2.0, Math.max(0.5, recentRate / campaign.conversions));
    } else if (recentConvs > 0) {
      // No 90-day conversions but some recent conversions - positive trend
      conversionTrend = 1.5; // Moderate boost
    } else {
      // No conversions in either period - slightly negative
      conversionTrend = 0.8; // Larger penalty for no conversions
    }
    
    // Calculate efficiency trend - compare recent vs. 90-day efficiency
    // (Get recent cost for the same period as recent conversions)
    const recentCost = getRecentCost(campaign.campaign, CONFIG.RECENT_PERFORMANCE_PERIOD);
    campaign.recentCost = recentCost; // Store for logging
    
    // Calculate efficiency ratios (share of conversions / share of cost)
    // These values will be available from calculateTrendFactorForAll after that function has run
    let efficiencyTrend = 1.0; // Default to neutral
    
    if (campaign.recentEfficiencyRatio && campaign.longTermEfficiencyRatio) {
      // If efficiency is improving, boost the trend factor
      if (campaign.recentEfficiencyRatio > campaign.longTermEfficiencyRatio) {
        efficiencyTrend = Math.min(1.5, Math.max(1.0, 
          campaign.recentEfficiencyRatio / Math.max(0.1, campaign.longTermEfficiencyRatio)));
      } 
      // If efficiency is declining, reduce the trend factor more aggressively
      else if (campaign.recentEfficiencyRatio < campaign.longTermEfficiencyRatio) {
        // Calculate the decline ratio (how much worse it's getting)
        const declineRatio = campaign.recentEfficiencyRatio / Math.max(0.1, campaign.longTermEfficiencyRatio);
        
        // Apply a more severe penalty for efficiency decline
        efficiencyTrend = Math.max(0.4, Math.min(1.0, declineRatio * 0.8));
      }
      
      // If recent efficiency ratio is very poor, apply an even stronger penalty
      if (campaign.recentEfficiencyRatio < 0.7) {
        efficiencyTrend = Math.min(efficiencyTrend, 0.7);
      }
      
      // If the campaign has BOTH declining efficiency AND was previously efficient,
      // apply an even stronger penalty (this targets previously good campaigns that are declining)
      if (campaign.recentEfficiencyRatio < campaign.longTermEfficiencyRatio && 
          campaign.longTermEfficiencyRatio > 1.0) {
        efficiencyTrend = Math.min(efficiencyTrend, 0.6);
      }
    }
    
    // Consider the impression share lost to budget - be more cautious with budget increases
    // when impression share lost is low
	let impressionShareFactor = 1.0;
	// For higher impression share loss, give higher factor (more budget)
	if (campaign.budgetImpressionShareLost > 50) {
	  impressionShareFactor = 1.0 + (Math.min(campaign.budgetImpressionShareLost, 80) - 50) / 100;
	} 
	// For lower impression share loss, reduce factor (less budget)
	else if (campaign.budgetImpressionShareLost < 30) {
	  impressionShareFactor = 0.8 + (campaign.budgetImpressionShareLost / 150);
	}
    
    // Combine factors with different weights
    // Give efficiency trend a higher weight to prioritize efficiency
    const combinedFactor = (conversionTrend * 0.4) + (efficiencyTrend * 0.4) + (impressionShareFactor * 0.2);
    
    // Limit the final trend factor to a reasonable range (0.5 to 2.0)
    const finalTrendFactor = Math.min(2.0, Math.max(0.5, combinedFactor));
    
    // Log the calculation components for transparency
    Logger.log(`Trend calculation for '${campaign.name}': 
      - Conversion trend: ${conversionTrend.toFixed(2)} (${CONFIG.RECENT_PERFORMANCE_PERIOD}-day: ${recentConvs.toFixed(2)}, 90-day: ${campaign.conversions.toFixed(2)})
      - Efficiency trend: ${efficiencyTrend.toFixed(2)} (Recent: ${campaign.recentEfficiencyRatio ? campaign.recentEfficiencyRatio.toFixed(2) : "N/A"}, 90-day: ${campaign.longTermEfficiencyRatio ? campaign.longTermEfficiencyRatio.toFixed(2) : "N/A"})
      - Impression share factor: ${impressionShareFactor.toFixed(2)} (IS Lost: ${campaign.budgetImpressionShareLost.toFixed(2)}%)
      - Final trend factor: ${finalTrendFactor.toFixed(2)}`);
    
    return finalTrendFactor;
  } catch (e) {
    Logger.log("Error calculating trend factor for campaign " + campaign.name + ": " + e);
    return 1.0; // Default to neutral
  }
}

function logCurrentBudgetStatus(campaignData) {
  try {
    if (!campaignData || !campaignData.campaigns) {
      Logger.log("No campaign data available for status logging");
      return;
    }

    Logger.log("\n===== CURRENT CAMPAIGN BUDGET STATUS =====");
    Logger.log(`Showing data for specific conversion action: "${CONFIG.CONVERSION_SETTINGS.CONVERSION_ACTION_NAME}"`);
    Logger.log("Campaign Name, Shared Budget Group, Optimization Objective, Daily Budget, Budget/Conv Ratio, 21-Day Convs, 90-Day Convs, Trend Factor, 90-Day Efficiency, Recent Efficiency");
    
    const totalConversions = campaignData.totalConversions || 0;
    
    for (const campaign of campaignData.campaigns) {
      try {
        const budgetConvRatio = totalConversions > 0 ? 
          (campaign.currentDailyBudget / totalConversions) : 0;
        
        const trendFactor = campaign.trendFactor || 1.0;
        
        // Fix: Use the correct property names
        const efficiency90Day = campaign.longTermEfficiencyRatio || 0;
        const recentEfficiency = campaign.recentEfficiencyRatio || 0;
        
        Logger.log(
          `${campaign.name}, ` +
          `${campaign.isSharedBudget ? 'Shared' : 'Individual'}, ` +
          `${campaign.optimizationObjective || 'N/A'}, ` +
          `$${campaign.currentDailyBudget.toFixed(2)}, ` +
          `$${budgetConvRatio.toFixed(2)}, ` +
          `${campaign.recentConversions || 0}, ` +
          `${campaign.conversions || 0}, ` +
          `${trendFactor.toFixed(2)}, ` +
          `${efficiency90Day.toFixed(2)}, ` +
          `${recentEfficiency.toFixed(2)}`
        );
      } catch (e) {
        Logger.log(`Error logging campaign ${campaign.name}: ${e}`);
        continue;
      }
    }
    
    Logger.log("===========================================\n");
  } catch (e) {
    Logger.log(`Error in logCurrentBudgetStatus: ${e}`);
  }
}

function processSharedBudgets(campaignData, dateRange) {
  Logger.log("Processing shared budget campaigns...");
  
  // Calculate today's maximum daily budget based on the monthly budget and days in current month
  const maxDailyBudget = calculateMaxDailyBudget();
  
  // Group campaigns by shared budget ID
  const sharedBudgets = {};
  
  campaignData.campaigns.forEach(campaign => {
    if (campaign.isSharedBudget && campaign.sharedBudgetId) {
      if (!sharedBudgets[campaign.sharedBudgetId]) {
        sharedBudgets[campaign.sharedBudgetId] = {
          campaigns: [],
          totalBudget: 0,
          totalConversions: 0,
          totalCost: 0,
          totalConversionValue: 0,
          totalClicks: 0,
          totalImpressions: 0,
          avgBudgetImpressionShareLost: 0,
          avgImpressionShare: 0,
          bidStrategies: new Set()
        };
      }
      
      const budgetGroup = sharedBudgets[campaign.sharedBudgetId];
      budgetGroup.campaigns.push(campaign);
      budgetGroup.totalBudget = campaign.currentDailyBudget; // All campaigns in the group have the same budget
      budgetGroup.totalConversions += campaign.conversions;
      budgetGroup.totalCost += campaign.cost;
      budgetGroup.totalConversionValue += campaign.conversionValue;
      budgetGroup.totalClicks += campaign.clicks;
      budgetGroup.totalImpressions += campaign.impressions;
      budgetGroup.bidStrategies.add(campaign.bidStrategy);
    }
  });
  
  // Process each shared budget group
  for (const budgetId in sharedBudgets) {
    const budget = sharedBudgets[budgetId];
    
    // Calculate averages and aggregated metrics
    budget.avgBudgetImpressionShareLost = budget.campaigns.reduce((sum, campaign) => 
      sum + campaign.budgetImpressionShareLost, 0) / budget.campaigns.length;
    
    budget.avgImpressionShare = budget.campaigns.reduce((sum, campaign) => 
      sum + campaign.impressionShare, 0) / budget.campaigns.length;
    
    // Calculate ROI, CPA, and ROAS for the budget group
    budget.roi = budget.totalCost > 0 ? budget.totalConversionValue / budget.totalCost : 0;
    budget.cpa = budget.totalConversions > 0 ? budget.totalCost / budget.totalConversions : 0;
    budget.roas = budget.totalCost > 0 ? (budget.totalConversionValue / budget.totalCost) * 100 : 0;
    
    // Calculate average trend factor for the shared budget group
    budget.avgTrendFactor = budget.campaigns.reduce((sum, campaign) => 
      sum + campaign.trendFactor, 0) / budget.campaigns.length;
    
    // NEW: Calculate average day-of-week multiplier for the shared budget group
    if (CONFIG.DAY_OF_WEEK.ENABLED) {
      let totalDowMultiplier = 0;
      let dowCampaignsWithData = 0;
      
      budget.campaigns.forEach(campaign => {
        if (campaign.dayOfWeekAdjustment && campaign.dayOfWeekAdjustment.confidence >= 0.3) {
          totalDowMultiplier += campaign.dayOfWeekAdjustment.appliedMultiplier;
          dowCampaignsWithData++;
        }
      });
      
      budget.avgDowMultiplier = dowCampaignsWithData > 0 ? 
        totalDowMultiplier / dowCampaignsWithData : 1.0;
      
      budget.dowCampaignsWithData = dowCampaignsWithData;
    } else {
      budget.avgDowMultiplier = 1.0;
    }
    
    // Get the shared budget object directly
    const sharedBudget = campaignData.sharedBudgetObjects[budgetId];
    
    if (sharedBudget) {
      const currentAmount = sharedBudget.getAmount();
      
      // Determine if budget adjustments are needed based on dominant bid strategy
      let newBudgetAmount = currentAmount;
      let adjustmentFactor = 1;
      let adjustmentReason = "";
      
      // Identify the dominant bid strategy (if there's a clear one)
      let dominantStrategy = null;
      if (budget.bidStrategies.size === 1) {
        dominantStrategy = Array.from(budget.bidStrategies)[0];
      } else {
        // Use ROI-based approach for mixed strategies
        dominantStrategy = 'MANUAL_CPC'; // Default to ROI-based
        adjustmentReason += "Mixed bid strategies, using ROI-based assessment. ";
      }
      
      // Apply strategy-specific logic
      const strategyConfig = CONFIG.STRATEGY_THRESHOLDS[dominantStrategy] || 
                            CONFIG.STRATEGY_THRESHOLDS['MANUAL_CPC'];
      
      let performanceMetric;
      let isPerformingWell;
      
      switch (dominantStrategy) {
        case 'TARGET_CPA':
          performanceMetric = budget.cpa;
          isPerformingWell = performanceMetric <= strategyConfig.threshold || strategyConfig.threshold === 0;
          break;
          
        case 'TARGET_ROAS':
          performanceMetric = budget.roas;
          isPerformingWell = performanceMetric >= strategyConfig.threshold;
          break;
            
        case 'MAXIMIZE_CONVERSIONS':
          performanceMetric = budget.totalConversions;
          isPerformingWell = performanceMetric > 0;
          break;
          
        case 'MAXIMIZE_CONVERSION_VALUE':
          performanceMetric = budget.totalConversionValue;
          isPerformingWell = performanceMetric > 0;
          break;
          
        case 'TARGET_IMPRESSION_SHARE':
          performanceMetric = budget.avgImpressionShare;
          isPerformingWell = performanceMetric >= strategyConfig.threshold;
          break;
          
        case 'MAXIMIZE_CLICKS':
          performanceMetric = budget.totalClicks;
          isPerformingWell = performanceMetric > 0;
          break;
          
        case 'MANUAL_CPC':
        default:
          performanceMetric = budget.roi;
          isPerformingWell = performanceMetric >= strategyConfig.threshold;
          break;
      }
      
      // Apply budget adjustments based on performance
      if (budget.avgBudgetImpressionShareLost > CONFIG.MAX_IMPRESSION_SHARE_LOST_TO_BUDGET && isPerformingWell) {
        // Increase budget proportionally to impression share lost
        const lossRatio = budget.avgBudgetImpressionShareLost / CONFIG.MAX_IMPRESSION_SHARE_LOST_TO_BUDGET;
        adjustmentFactor = Math.min(1 + (CONFIG.MAX_ADJUSTMENT_PERCENTAGE / 100 * lossRatio), 
                                   1 + (CONFIG.MAX_ADJUSTMENT_PERCENTAGE / 100));
        
        adjustmentReason += "Increasing budget due to high impression share loss (" + 
                          budget.avgBudgetImpressionShareLost.toFixed(2) + "%) with good performance.";
      } 
      // If not losing impression share due to budget and not performing well, decrease budget
      else if (budget.avgBudgetImpressionShareLost < CONFIG.MAX_IMPRESSION_SHARE_LOST_TO_BUDGET / 2 && !isPerformingWell) {
        adjustmentFactor = 1 - (CONFIG.MAX_ADJUSTMENT_PERCENTAGE / 100);
        
        adjustmentReason += "Decreasing budget due to low impression share loss (" + 
                          budget.avgBudgetImpressionShareLost.toFixed(2) + "%) with suboptimal performance.";
      } else {
        adjustmentReason += "No adjustment needed. Impression share loss: " + 
                          budget.avgBudgetImpressionShareLost.toFixed(2) + "%, Performance metric: " + 
                          performanceMetric.toFixed(2);
      }
      
      // Apply trend factor to adjustment
      adjustmentFactor = adjustmentFactor * budget.avgTrendFactor;
      adjustmentReason += ` Applied trend factor: ${budget.avgTrendFactor.toFixed(2)} (${CONFIG.RECENT_PERFORMANCE_PERIOD}-day vs ${CONFIG.LOOKBACK_PERIOD}-day performance).`;
      
      // NEW: Apply day-of-week adjustment if enabled
      if (CONFIG.DAY_OF_WEEK.ENABLED && budget.dowCampaignsWithData > 0) {
        const dowInfluence = Math.min(0.8, budget.dowCampaignsWithData / budget.campaigns.length) * 
                          CONFIG.DAY_OF_WEEK.INFLUENCE_WEIGHT;
        
        // Blend the trend-based adjustment with day-of-week factor
        const blendedAdjustment = (adjustmentFactor * (1 - dowInfluence)) + 
                                (budget.avgDowMultiplier * dowInfluence);
        
        adjustmentReason += ` Applied day-of-week adjustment: ${budget.avgDowMultiplier.toFixed(2)} with ${(dowInfluence * 100).toFixed(1)}% influence (${budget.dowCampaignsWithData}/${budget.campaigns.length} campaigns with data).`;
        
        // Update the adjustment factor
        adjustmentFactor = blendedAdjustment;
      }
      
      // Calculate new budget amount
      newBudgetAmount = currentAmount * adjustmentFactor;
      
      // Ensure we don't exceed the maximum daily budget for shared budgets
      if (newBudgetAmount > maxDailyBudget) {
        newBudgetAmount = maxDailyBudget;
        adjustmentReason += ` Budget capped at daily maximum of ${maxDailyBudget.toFixed(2)}.`;
      }
      
      // Calculate the percentage change
      const percentChange = ((newBudgetAmount / currentAmount - 1) * 100).toFixed(2);
      
      // Always log the budget change, regardless of size
      Logger.log("Budget for shared budget group '" + budgetId + "': " + 
                currentAmount.toFixed(2) + " to " + newBudgetAmount.toFixed(2) + 
                " (" + percentChange + "% change)");
      Logger.log("Reason: " + adjustmentReason);
      
      // Apply the new budget if it's different enough
      if (Math.abs(newBudgetAmount - currentAmount) / currentAmount > (CONFIG.MIN_ADJUSTMENT_PERCENTAGE / 100)) {
        if (CONFIG.PREVIEW_MODE) {
          Logger.log("  [PREVIEW MODE: Budget would be updated from " + currentAmount.toFixed(2) + 
                   " to " + newBudgetAmount.toFixed(2) + " (" + percentChange + "% change)]");
        } else {
          sharedBudget.setAmount(newBudgetAmount);
          Logger.log("  [Budget updated]");
        }
        budget.newAmount = newBudgetAmount;
      } else {
        Logger.log("  [Change below minimum threshold of " + CONFIG.MIN_ADJUSTMENT_PERCENTAGE + "% - not applied]");
        budget.newAmount = currentAmount;
      }
    } else {
      Logger.log("Could not find budget object for shared budget group '" + budgetId + "'");
    }
  }
  
  // Store shared budgets data for summary
  campaignData.processedSharedBudgets = sharedBudgets;
}

/**
 * Analyzes future days in the current month to determine optimal budget allocation
 * @param {Object} campaign - The campaign to analyze
 * @param {Object} dayPerformance - Current day-of-week performance data
 * @param {Object} dayIndices - Day-of-week performance indices
 * @param {Object} pacingInfo - Budget pacing information
 * @return {Object} Analysis results including future days info and recommended adjustments
 */
function analyzeFutureDays(campaign, dayIndices, pacingInfo) {
  try {
    const today = new Date();
    const accountTimezone = AdsApp.currentAccount().getTimeZone();
    const currentDay = parseInt(Utilities.formatDate(today, accountTimezone, "u")) % 7;
    const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
    const currentDayOfMonth = today.getDate();
    const daysRemaining = daysInMonth - currentDayOfMonth;
    
    // Get today's performance index
    const todayIndex = dayIndices && dayIndices.dayIndices && dayIndices.dayIndices[currentDay] ? 
      dayIndices.dayIndices[currentDay].conversionIndex : 1;
    
    // Analyze future days
    const futureDays = [];
    let highPerformanceDaysAhead = 0;
    let totalFutureIndex = 0;
    
    // Track unique days of week simulated
    const simulatedDays = {};
    
    for (let i = 1; i <= daysRemaining; i++) {
      const futureDate = new Date(today);
      futureDate.setDate(today.getDate() + i);
      const futureDayOfWeek = futureDate.getDay();
      
      // Get day performance data
      const dayData = dayIndices && dayIndices.dayIndices && dayIndices.dayIndices[futureDayOfWeek];
      
      // Validate pattern if we have data
      let patternValidation = {
        isSignificant: false,
        patternStrength: 0,
        reason: "Insufficient data"
      };
      
      if (dayData && dayData.sampleSize >= CONFIG.DAY_OF_WEEK.MIN_DATA_POINTS) {
        // Calculate pattern significance
        const indices = Object.values(dayIndices.dayIndices)
          .filter(d => d.sampleSize >= CONFIG.DAY_OF_WEEK.MIN_DATA_POINTS)
          .map(d => d.conversionIndex);
        
        if (indices.length >= 3) {
          const mean = indices.reduce((sum, val) => sum + val, 0) / indices.length;
          const variance = indices.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / indices.length;
          const stdDev = Math.sqrt(variance);
          const patternStrength = stdDev / mean;
          
          patternValidation = {
            isSignificant: patternStrength >= 0.1 && indices.length >= 3,
            patternStrength: patternStrength,
            stdDev: stdDev,
            mean: mean,
            reason: patternStrength >= 0.1 ? 
              `Pattern is significant (strength: ${patternStrength.toFixed(2)})` :
              `Pattern is not significant (strength: ${patternStrength.toFixed(2)})`
          };
        }
      }
      
      // Get day index with pattern validation
      const dayIndex = dayData ? dayData.conversionIndex : 1;
      const confidence = dayData ? dayData.confidence : 0;
      
      // Only count as high performance if pattern is significant
      if (dayIndex > 1.1 && patternValidation.isSignificant) {
        highPerformanceDaysAhead++;
      }
      
      totalFutureIndex += dayIndex;
      
      // Track this day of week
      simulatedDays[futureDayOfWeek] = true;
      
      futureDays.push({
        date: futureDate,
        dayOfWeek: futureDayOfWeek,
        performanceIndex: dayIndex,
        confidence: confidence
      });
      
      // Check for special events
      let specialEvent = null;
      for (const event of CONFIG.DAY_OF_WEEK.SPECIAL_EVENTS) {
        if (event.date === dateString) {
          specialEvent = event;
          break;
        }
      }
      
      // Update pattern validation with current data
      patternValidation = validateDayOfWeekPattern(dayIndices.dayIndices, dayIndex.conversionIndex);
      
      // Calculate expected performance
      let expectedMultiplier = 1.0;
      let reason = "No significant pattern";
      
      if (specialEvent) {
        expectedMultiplier = specialEvent.multiplier;
        confidence = 1.0;
        reason = `Special event: ${specialEvent.name}`;
      } else if (dayData && dayData.sampleSize >= CONFIG.DAY_OF_WEEK.MIN_DATA_POINTS) {
        if (patternValidation.isSignificant) {
          expectedMultiplier = dayData.conversionIndex;
          confidence = dayData.confidence * patternValidation.patternStrength;
          reason = `Day-of-week pattern (strength: ${patternValidation.patternStrength.toFixed(2)})`;
        }
      }
      
      // Calculate budget adjustment strategy
      let budgetStrategy = "neutral";
      if (confidence >= 0.7) {
        if (expectedMultiplier > 1.2) {
          budgetStrategy = "aggressive_increase";
        } else if (expectedMultiplier > 1.1) {
          budgetStrategy = "moderate_increase";
        } else if (expectedMultiplier < 0.8) {
          budgetStrategy = "aggressive_decrease";
        } else if (expectedMultiplier < 0.9) {
          budgetStrategy = "moderate_decrease";
        }
      }
      
      // Update the last added day with additional information
      const lastDay = futureDays[futureDays.length - 1];
      lastDay.expectedMultiplier = expectedMultiplier;
      lastDay.reason = reason;
      lastDay.budgetStrategy = budgetStrategy;
      lastDay.isSpecialEvent = !!specialEvent;
      lastDay.specialEventName = specialEvent ? specialEvent.name : null;
      lastDay.patternValidation = patternValidation;
    }
    
    // Log the analysis
    Logger.log(`\nFuture Days Analysis for ${campaign.getName()}:`);
    Logger.log("----------------------------------------");
    
    // Group days by budget strategy
    const strategyGroups = {};
    futureDays.forEach(day => {
      if (!strategyGroups[day.budgetStrategy]) {
        strategyGroups[day.budgetStrategy] = [];
      }
      strategyGroups[day.budgetStrategy].push(day);
    });
    
    // Log each strategy group
    Object.entries(strategyGroups).forEach(([strategy, days]) => {
      Logger.log(`\n${strategy.replace(/_/g, ' ').toUpperCase()} Days:`);
      days.forEach(day => {
        Logger.log(`  ${day.date} (${day.dayName}):`);
        Logger.log(`    Expected Multiplier: ${day.expectedMultiplier.toFixed(2)}`);
        Logger.log(`    Confidence: ${(day.confidence * 100).toFixed(1)}%`);
        Logger.log(`    Reason: ${day.reason}`);
        if (day.isSpecialEvent) {
          Logger.log(`    Special Event: ${day.specialEventName}`);
        }
        if (day.patternValidation.isSignificant) {
          Logger.log(`    Pattern Strength: ${day.patternValidation.patternStrength.toFixed(2)}`);
        }
      });
    });
    
    return futureDays;
  } catch (e) {
    Logger.log(`Error analyzing future days: ${e}`);
    return [];
  }
}

/**
 * Redistributes budgets between campaigns based on performance and pacing
 * @param {Array} campaigns - Array of campaign objects
 * @param {Object} pacingInfo - Budget pacing information
 */
function redistributeBudgets(campaigns, pacingInfo) {
  try {
    // Safety checks for redistribution
    if (!campaigns || campaigns.length < 2) {
      Logger.log("Not enough campaigns for redistribution");
      return;
    }
    
    // Check if we're over pace
    if (pacingInfo && pacingInfo.paceStatus === 'OVER_PACE') {
      Logger.log("Skipping redistribution - account is over pace");
      return;
    }
    
    // Check if we have enough days remaining for meaningful redistribution
    if (pacingInfo && pacingInfo.daysRemaining < 3) {
      Logger.log("Skipping redistribution - less than 3 days remaining");
      return;
    }
    
    // Check if we have too many campaigns needing redistribution
    const campaignsNeedingIncrease = campaigns.filter(c => c.needsBudgetRedistribution);
    if (campaignsNeedingIncrease.length > CONFIG.MEMORY_LIMITS.MAX_CAMPAIGNS_PER_BATCH) {
      Logger.log(`Warning: Too many campaigns needing redistribution (${campaignsNeedingIncrease.length})`);
      return;
    }
    
    // Calculate total budget available for redistribution
    let totalBudgetAvailable = 0;
    let errorCount = 0;
    
    for (const campaign of campaigns) {
      try {
        if (!campaign.needsBudgetRedistribution) {
          const currentBudget = campaign.currentDailyBudget;
          const maxAllowedBudget = pacingInfo ? 
            currentBudget * pacingInfo.maxDailyBudgetMultiplier : 
            currentBudget * CONFIG.MAX_DAILY_BUDGET_MULTIPLIER;
          
          if (currentBudget < maxAllowedBudget) {
            totalBudgetAvailable += maxAllowedBudget - currentBudget;
          }
        }
      } catch (e) {
        errorCount++;
        Logger.log(`Error calculating available budget for ${campaign.name}: ${e}`);
        continue;
      }
    }
    
    // Safety check for total available budget
    if (totalBudgetAvailable <= 0) {
      Logger.log("No budget available for redistribution");
      return;
    }
    
    // Calculate performance scores for campaigns needing increases
    const performanceScores = [];
    for (const campaign of campaignsNeedingIncrease) {
      try {
        const score = calculatePerformanceScore(campaign);
        if (score > 0) {
          performanceScores.push({
            campaign: campaign,
            score: score
          });
        }
      } catch (e) {
        Logger.log(`Error calculating performance score for ${campaign.name}: ${e}`);
        continue;
      }
    }
    
    // Safety check for valid performance scores
    if (performanceScores.length === 0) {
      Logger.log("No valid performance scores for redistribution");
      return;
    }
    
    // Sort by performance score
    performanceScores.sort((a, b) => b.score - a.score);
    
    // Calculate total performance score
    const totalScore = performanceScores.reduce((sum, ps) => sum + ps.score, 0);
    
    // Distribute budget based on performance scores
    let redistributedBudget = 0;
    let appliedRedistributions = 0;
    
    for (const {campaign, score} of performanceScores) {
      try {
        const currentBudget = campaign.currentDailyBudget;
        const maxAllowedBudget = pacingInfo ? 
          currentBudget * pacingInfo.maxDailyBudgetMultiplier : 
          currentBudget * CONFIG.MAX_DAILY_BUDGET_MULTIPLIER;
        
        // Calculate budget increase based on performance score
        const budgetIncrease = (score / totalScore) * totalBudgetAvailable;
        const newBudget = Math.min(currentBudget + budgetIncrease, maxAllowedBudget);
        
        // Log redistribution details
        Logger.log(`\nRedistribution for ${campaign.name}:`);
        Logger.log(`  - Current budget: ${currentBudget.toFixed(2)}`);
        Logger.log(`  - Performance score: ${score.toFixed(2)}`);
        Logger.log(`  - Budget increase: ${budgetIncrease.toFixed(2)}`);
        Logger.log(`  - New budget: ${newBudget.toFixed(2)}`);
        Logger.log(`  - Maximum allowed: ${maxAllowedBudget.toFixed(2)}`);
        
        // Apply the budget change if it's significant enough
        if (Math.abs(newBudget - currentBudget) / currentBudget > (CONFIG.MIN_ADJUSTMENT_PERCENTAGE / 100)) {
          if (CONFIG.PREVIEW_MODE) {
            Logger.log(`  [PREVIEW MODE: Budget would be updated from ${currentBudget.toFixed(2)} to ${newBudget.toFixed(2)}]`);
          } else {
            campaign.campaign.getBudget().setAmount(newBudget);
            Logger.log(`  [Budget updated]`);
            redistributedBudget += newBudget - currentBudget;
            appliedRedistributions++;
          }
        } else {
          Logger.log(`  [Change below minimum threshold of ${CONFIG.MIN_ADJUSTMENT_PERCENTAGE}% - not applied]`);
        }
      } catch (e) {
        Logger.log(`Error applying redistribution for ${campaign.name}: ${e}`);
        continue;
      }
    }
    
    // Log redistribution summary
    Logger.log(`\nBudget Redistribution Summary:
      - Total budget available: ${totalBudgetAvailable.toFixed(2)}
      - Budget redistributed: ${redistributedBudget.toFixed(2)}
      - Redistributions applied: ${appliedRedistributions}
      - Errors encountered: ${errorCount}
      - Success rate: ${((appliedRedistributions / (appliedRedistributions + errorCount)) * 100).toFixed(1)}%`);
    
  } catch (e) {
    Logger.log(`Critical error in redistributeBudgets: ${e}`);
    throw e;
  }
}

function calculateFuturePerformanceScore(campaign) {
  let score = 1.0;
  
  // Factor 1: Future Day Performance (40%)
  if (campaign.futureAnalysis && campaign.futureAnalysis.futureDays) {
    const futureDays = campaign.futureAnalysis.futureDays;
    const avgFutureIndex = futureDays.reduce((sum, day) => sum + day.performanceIndex, 0) / futureDays.length;
    score *= (1 + ((avgFutureIndex - 1) * 0.4));
  }
  
  // Factor 2: Pattern Significance (30%)
  if (campaign.futureAnalysis && campaign.futureAnalysis.patternValidation) {
    const pattern = campaign.futureAnalysis.patternValidation;
    if (pattern.isSignificant) {
      score *= (1 + (pattern.patternStrength * 0.3));
    }
  }
  
  // Factor 3: Recent Performance (30%)
  if (campaign.trendFactor) {
    score *= (1 + ((campaign.trendFactor - 1) * 0.3));
  }
  
  return score;
}

function processCampaignBudgets(campaignData, dateRange, pacingInfo) {
  try {
    // Log budget pacing impact
    Logger.log("\n===== BUDGET PACING IMPACT ON CAMPAIGNS =====");
    if (pacingInfo) {
      const spentToDate = pacingInfo.totalSpend || 0;
      const monthlyBudget = pacingInfo.monthlyBudget || 1;
      const remainingBudget = pacingInfo.remainingBudget || 0;
      const maxDailyMultiplier = pacingInfo.maxDailyBudgetMultiplier || 1;
      const idealDailyRemaining = pacingInfo.idealDailyRemaining || 0;
      const spendPercentage = (spentToDate / monthlyBudget * 100).toFixed(1);
      
      Logger.log(`Current pace status: ${pacingInfo.paceStatus || 'unknown'}`);
      Logger.log(`Monthly spend to date: ${spentToDate.toFixed(2)} (${spendPercentage}%)`);
      Logger.log(`Remaining budget: ${remainingBudget.toFixed(2)}`);
      Logger.log(`Maximum daily budget multiplier: ${maxDailyMultiplier.toFixed(2)}`);
      Logger.log(`Ideal daily remaining: ${idealDailyRemaining.toFixed(2)}`);
    } else {
      Logger.log("Budget pacing information not available");
    }
    
    // Calculate initial total daily budget
    const initialTotalBudget = campaignData.campaigns.reduce((sum, c) => sum + c.currentDailyBudget, 0);
    
    // First pass: Calculate proposed adjustments with safety checks
    const proposedAdjustments = [];
    let processedCampaigns = 0;
    let errorCount = 0;
    
    for (const campaign of campaignData.campaigns) {
      // Skip shared budget campaigns as they're handled separately
      if (campaign.isSharedBudget) continue;
      
      // Safety check: Limit number of campaigns processed
      if (processedCampaigns >= CONFIG.MEMORY_LIMITS.MAX_CAMPAIGNS_PER_BATCH) {
        Logger.log("Warning: Reached maximum campaigns per batch limit");
        break;
      }
      
      try {
        // Calculate adjustment with all available factors
        const adjustment = safeCalculateAdjustment(campaign, {
          trendFactor: campaign.trendFactor,
          dayOfWeekAdjustment: campaign.dayOfWeekAdjustment,
          performanceScore: campaign.performanceScore,
          efficiencyMetrics: campaign.efficiencyMetrics
        });
        
        // Validate adjustment values
        if (!adjustment || !adjustment.adjustmentFactor || isNaN(adjustment.adjustmentFactor)) {
          Logger.log(`Warning: Invalid adjustment calculated for campaign ${campaign.name}`);
          continue;
        }
        
        // Calculate proposed new budget
        const currentBudget = campaign.currentDailyBudget;
        const proposedBudget = currentBudget * adjustment.adjustmentFactor;
        
        // Store the proposed adjustment
        proposedAdjustments.push({
          campaign: campaign,
          currentBudget: currentBudget,
          proposedBudget: proposedBudget,
          adjustmentFactor: adjustment.adjustmentFactor,
          adjustmentReason: adjustment.reason,
          confidence: adjustment.confidence,
          performanceScore: campaign.performanceScore
        });
        
        processedCampaigns++;
      } catch (e) {
        errorCount++;
        Logger.log(`Error processing campaign ${campaign.name}: ${e}`);
        continue;
      }
    }
    
    // Log processing summary
    Logger.log(`\nCampaign Processing Summary:
      - Total campaigns processed: ${processedCampaigns}
      - Errors encountered: ${errorCount}
      - Success rate: ${((processedCampaigns / (processedCampaigns + errorCount)) * 100).toFixed(1)}%`);
    
    // Sort adjustments by performance score (highest to lowest)
    proposedAdjustments.sort((a, b) => b.performanceScore - a.performanceScore);
    
    // Calculate total proposed budget
    const totalProposedBudget = proposedAdjustments.reduce((sum, adj) => sum + adj.proposedBudget, 0);
    
    // Calculate the budget adjustment needed to stay within pacing limits
    const maxAllowedBudget = pacingInfo ? pacingInfo.idealDailyRemaining : initialTotalBudget;
    const budgetAdjustmentRatio = maxAllowedBudget / totalProposedBudget;
    
    // Second pass: Apply final adjustments with budget balancing
    let appliedAdjustments = 0;
    let skippedAdjustments = 0;
    
    for (const adjustment of proposedAdjustments) {
      try {
        const campaign = adjustment.campaign;
        const currentBudget = adjustment.currentBudget;
        
        // Calculate final budget with balancing
        let finalBudget = adjustment.proposedBudget * budgetAdjustmentRatio;
        
        // Ensure we don't exceed the maximum daily budget multiplier
        if (pacingInfo && pacingInfo.maxDailyBudgetMultiplier) {
          const maxAllowedForCampaign = currentBudget * pacingInfo.maxDailyBudgetMultiplier;
          finalBudget = Math.min(finalBudget, maxAllowedForCampaign);
        }
        
        // Log campaign-specific adjustments
        Logger.log(`\nAdjustments for ${campaign.name}:`);
        Logger.log(`  - Current budget: ${currentBudget.toFixed(2)}`);
        Logger.log(`  - Proposed adjustment: ${adjustment.adjustmentFactor.toFixed(2)}`);
        Logger.log(`  - Budget balancing ratio: ${budgetAdjustmentRatio.toFixed(2)}`);
        Logger.log(`  - Final budget: ${finalBudget.toFixed(2)}`);
        Logger.log(`  - Confidence: ${(adjustment.confidence * 100).toFixed(1)}%`);
        Logger.log(`  - Reason: ${adjustment.adjustmentReason}`);
        
        // Store the new budget in the campaign object
        campaign.newBudget = finalBudget;
        
        // Apply the budget change if it's significant enough
        if (Math.abs(finalBudget - currentBudget) / currentBudget > (CONFIG.MIN_ADJUSTMENT_PERCENTAGE / 100)) {
          if (CONFIG.PREVIEW_MODE) {
            Logger.log(`  [PREVIEW MODE: Budget would be updated from ${currentBudget.toFixed(2)} to ${finalBudget.toFixed(2)}]`);
          } else {
            const budget = campaign.campaign.getBudget();
            if (budget) {
              budget.setAmount(finalBudget);
              Logger.log(`  [Budget updated]`);
              appliedAdjustments++;
            } else {
              Logger.log(`  [Warning: Could not get budget object for campaign ${campaign.name}]`);
              skippedAdjustments++;
            }
          }
        } else {
          Logger.log(`  [Change below minimum threshold of ${CONFIG.MIN_ADJUSTMENT_PERCENTAGE}% - not applied]`);
          campaign.newBudget = currentBudget; // Reset to current if change is too small
        }
      } catch (e) {
        Logger.log(`Error applying budget adjustment for ${adjustment.campaign.name}: ${e}`);
        continue;
      }
    }
    
    // After processing all campaigns, handle budget redistribution if needed
    if (campaignData.campaigns.some(c => c.needsBudgetRedistribution)) {
      try {
        redistributeBudgets(campaignData.campaigns, pacingInfo);
      } catch (e) {
        Logger.log(`Error during budget redistribution: ${e}`);
      }
    }
    
    Logger.log("===========================================\n");
  } catch (e) {
    Logger.log(`Critical error in processCampaignBudgets: ${e}`);
    throw e;
  }
}

// Helper function to calculate a performance score for budget balancing
function calculatePerformanceScore(campaign) {
  // Base score starts at 1.0
  let score = 1.0;
  
  // Factor 1: Conversion Efficiency (40% weight)
  const conversionEfficiency = campaign.conversions > 0 ? 
    campaign.conversionValue / campaign.cost : 0;
  score *= (1 + (conversionEfficiency * 0.4));
  
  // Factor 2: Recent Performance Trend (30% weight)
  if (campaign.trendFactor) {
    score *= (1 + ((campaign.trendFactor - 1) * 0.3));
  }
  
  // Factor 3: Impression Share Lost (20% weight)
  const impressionShareImpact = Math.max(0, 1 - (campaign.budgetImpressionShareLost / 100));
  score *= (1 + (impressionShareImpact * 0.2));
  
  // Factor 4: Day-of-Week Performance (10% weight)
  if (campaign.dayOfWeekAdjustment && campaign.dayOfWeekAdjustment.confidence >= 0.3) {
    const dowMultiplier = campaign.dayOfWeekAdjustment.appliedMultiplier;
    score *= (1 + ((dowMultiplier - 1) * 0.1));
  }
  
  // Apply confidence-based dampening
  let confidence = 1.0;
  
  // Reduce confidence if we have limited data
  if (campaign.conversions < 10) {
    confidence *= 0.8;
  }
  
  // Reduce confidence if impression share data is unreliable
  if (campaign.impressionShare < 0.1) {
    confidence *= 0.9;
  }
  
  // Apply confidence dampening
  score = 1 + ((score - 1) * confidence);
  
  // Ensure score stays within reasonable bounds
  score = Math.max(0.5, Math.min(2.0, score));
  
  return score;
}

function logSummary(campaignData) {
  const today = new Date();
  const dayOfWeek = parseInt(Utilities.formatDate(today, scriptTimezone, "u")) % 7;
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  Logger.log("\n===== EXECUTION SUMMARY =====");
  Logger.log("Date: " + today.toISOString());
  Logger.log("Today is: " + dayNames[dayOfWeek]);
  Logger.log("Total campaigns analyzed: " + campaignData.campaigns.length);
  Logger.log(`Using ${CONFIG.RECENT_PERFORMANCE_PERIOD}-day trending against ${CONFIG.LOOKBACK_PERIOD}-day baseline`);
  
  // Log day-specific adjustments summary
  if (CONFIG.DAY_OF_WEEK.ENABLED) {
    Logger.log("\n=== Day-of-Week Adjustments ===");
    const campaignsWithDayAdjustments = campaignData.campaigns.filter(c => c.dayOfWeekAdjustment);
    Logger.log(`Campaigns with day-specific adjustments: ${campaignsWithDayAdjustments.length}`);
    
    // Calculate average confidence and influence
    const avgConfidence = campaignsWithDayAdjustments.reduce((sum, c) => 
      sum + (c.dayOfWeekAdjustment.confidence || 0), 0) / campaignsWithDayAdjustments.length;
    
    Logger.log(`Average confidence score: ${(avgConfidence * 100).toFixed(1)}%`);
    
    // Group adjustments by size
    const adjustmentGroups = {
      significant: [], // >10% change
      moderate: [], // 5-10% change
      minor: [], // <5% change
      ignored: [] // Below confidence threshold
    };
    
    campaignsWithDayAdjustments.forEach(campaign => {
      const adj = campaign.dayOfWeekAdjustment;
      if (adj.confidence < 0.3) {
        adjustmentGroups.ignored.push(campaign);
      } else {
        const changePercent = Math.abs((adj.appliedMultiplier - 1) * 100);
        if (changePercent > 10) {
          adjustmentGroups.significant.push(campaign);
        } else if (changePercent > 5) {
          adjustmentGroups.moderate.push(campaign);
        } else {
          adjustmentGroups.minor.push(campaign);
        }
      }
    });
    
    Logger.log("\nAdjustment Distribution:");
    Logger.log(`  Significant adjustments (>10%): ${adjustmentGroups.significant.length} campaigns`);
    Logger.log(`  Moderate adjustments (5-10%): ${adjustmentGroups.moderate.length} campaigns`);
    Logger.log(`  Minor adjustments (<5%): ${adjustmentGroups.minor.length} campaigns`);
    Logger.log(`  Ignored (low confidence): ${adjustmentGroups.ignored.length} campaigns`);
    
    // Log special events if any
    const specialEvents = campaignsWithDayAdjustments.filter(c => c.dayOfWeekAdjustment.isSpecialEvent);
    if (specialEvents.length > 0) {
      Logger.log("\nSpecial Events Applied:");
      specialEvents.forEach(campaign => {
        Logger.log(`  ${campaign.name}: ${campaign.dayOfWeekAdjustment.specialEventName} (75% weight applied)`);
      });
    }
  }
  
  // Calculate budget totals
  let totalBudgetBefore = campaignData.totalBudget || 0;
  
  // Get individual campaign budget totals after adjustments
  const individualCampaigns = campaignData.campaigns.filter(c => !c.isSharedBudget);
  const individualBudgetAfter = individualCampaigns.reduce((sum, c) => sum + (c.newBudget || c.currentDailyBudget), 0);
  
  // Calculate shared budget totals after adjustments
  let sharedBudgetBefore = totalBudgetBefore - campaignData.totalIndividualBudget;
  let sharedBudgetAfter = 0;
  
  if (campaignData.processedSharedBudgets) {
    for (const budgetId in campaignData.processedSharedBudgets) {
      const budget = campaignData.processedSharedBudgets[budgetId];
      sharedBudgetAfter += budget.newAmount || budget.totalBudget;
    }
  } else {
    sharedBudgetAfter = sharedBudgetBefore; // No changes if not processed
  }
  
  // Calculate total budget after
  const totalBudgetAfter = individualBudgetAfter + sharedBudgetAfter;
  
  Logger.log("\n=== Budget Changes Summary ===");
  Logger.log(`Total budget: ${totalBudgetBefore.toFixed(2)} → ${totalBudgetAfter.toFixed(2)} (${((totalBudgetAfter/totalBudgetBefore - 1) * 100).toFixed(2)}% change)`);
  Logger.log(`Individual campaigns: ${campaignData.totalIndividualBudget.toFixed(2)} → ${individualBudgetAfter.toFixed(2)} (${((individualBudgetAfter/campaignData.totalIndividualBudget - 1) * 100).toFixed(2)}% change)`);
  Logger.log(`Shared budgets: ${sharedBudgetBefore.toFixed(2)} → ${sharedBudgetAfter.toFixed(2)} (${((sharedBudgetAfter/sharedBudgetBefore - 1) * 100).toFixed(2)}% change)`);
  
  // Log breakdown of changes by adjustment type
  const changes = {
    increased: [],
    decreased: [],
    unchanged: []
  };
  
  // Process individual campaigns
  individualCampaigns.forEach(campaign => {
    const oldBudget = campaign.currentDailyBudget;
    const newBudget = campaign.newBudget || oldBudget;
    const percentChange = (newBudget / oldBudget - 1) * 100;
    
    if (Math.abs(percentChange) < CONFIG.MIN_ADJUSTMENT_PERCENTAGE) {
      changes.unchanged.push({campaign, percentChange});
    } else if (percentChange > 0) {
      changes.increased.push({campaign, percentChange});
    } else {
      changes.decreased.push({campaign, percentChange});
    }
  });
  
  // Process shared budgets
  if (campaignData.processedSharedBudgets) {
    for (const budgetId in campaignData.processedSharedBudgets) {
      const budget = campaignData.processedSharedBudgets[budgetId];
      const oldBudget = budget.totalBudget;
      const newBudget = budget.newAmount || oldBudget;
      const percentChange = (newBudget / oldBudget - 1) * 100;
      
      if (Math.abs(percentChange) < CONFIG.MIN_ADJUSTMENT_PERCENTAGE) {
        changes.unchanged.push({campaign: {name: `Shared Budget Group ${budgetId}`}, percentChange});
      } else if (percentChange > 0) {
        changes.increased.push({campaign: {name: `Shared Budget Group ${budgetId}`}, percentChange});
      } else {
        changes.decreased.push({campaign: {name: `Shared Budget Group ${budgetId}`}, percentChange});
      }
    }
  }
  
  Logger.log("\nChange Distribution:");
  Logger.log(`  Increased budgets: ${changes.increased.length} campaigns`);
  Logger.log(`  Decreased budgets: ${changes.decreased.length} campaigns`);
  Logger.log(`  Unchanged budgets: ${changes.unchanged.length} campaigns`);
  
  // Log top changes
  if (changes.increased.length > 0) {
    changes.increased.sort((a, b) => b.percentChange - a.percentChange);
    Logger.log("\nLargest Increases:");
    changes.increased.slice(0, 5).forEach(({campaign, percentChange}) => {
      Logger.log(`  ${campaign.name}: +${percentChange.toFixed(2)}%`);
    });
  }
  
  if (changes.decreased.length > 0) {
    changes.decreased.sort((a, b) => a.percentChange - b.percentChange);
    Logger.log("\nLargest Decreases:");
    changes.decreased.slice(0, 5).forEach(({campaign, percentChange}) => {
      Logger.log(`  ${campaign.name}: ${percentChange.toFixed(2)}%`);
    });
  }
  
  // Log performance metrics
  Logger.log("\n=== Performance Metrics ===");
  const performanceData = {
    conversions: 0,
    cost: 0,
    conversionValue: 0
  };
  
  // Include both individual and shared budget campaigns
  campaignData.campaigns.forEach(campaign => {
    performanceData.conversions += campaign.conversions || 0;
    performanceData.cost += campaign.cost || 0;
    performanceData.conversionValue += campaign.conversionValue || 0;
  });
  
  Logger.log(`Total Conversions: ${performanceData.conversions}`);
  Logger.log(`Total Cost: ${performanceData.cost.toFixed(2)}`);
  Logger.log(`Total Conversion Value: ${performanceData.conversionValue.toFixed(2)}`);
  if (performanceData.cost > 0) {
    Logger.log(`Overall ROAS: ${(performanceData.conversionValue / performanceData.cost).toFixed(2)}`);
    Logger.log(`Cost per Conversion: ${(performanceData.cost / performanceData.conversions).toFixed(2)}`);
  }
  
  Logger.log("\n===== END SUMMARY =====\n");
}

function getRecentConversions(campaign, days, conversionActionName = null) {
  try {
    // Calculate date range for the specified number of days
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - days);
    
    const dateRange = {
      start: Utilities.formatDate(startDate, 'UTC', 'yyyyMMdd'),
      end: Utilities.formatDate(endDate, 'UTC', 'yyyyMMdd')
    };
    
    // If we're looking for a specific conversion action
    if (conversionActionName || CONFIG.CONVERSION_SETTINGS.USE_SPECIFIC_CONVERSION_ACTION) {
      const actionName = conversionActionName || CONFIG.CONVERSION_SETTINGS.CONVERSION_ACTION_NAME;
      
      return getSpecificConversionMetrics(campaign, dateRange, actionName).conversions;
    } else {
      // Use default conversions from the campaign stats
      const stats = campaign.getStatsFor(dateRange.start, dateRange.end);
      return stats.getConversions();
    }
  } catch (e) {
    Logger.log("Error getting recent conversion metrics for campaign " + campaign.getName() + ": " + e);
    return 0;
  }
}

function getCorrect90DayConversions(campaign, conversionActionName) {
  // This function uses our proven chunked approach directly
  // Calculate date range for 90 days
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 90);
  
  const dateRange = {
    start: Utilities.formatDate(startDate, 'UTC', 'yyyyMMdd'),
    end: Utilities.formatDate(endDate, 'UTC', 'yyyyMMdd')
  };
  
  // Use our chunked method that has been proven to work
  return getSpecificConversionMetrics(
    campaign, 
    dateRange, 
    conversionActionName
  ).conversions;
}

function testConversionActionAvailability() {
  Logger.log("\n===== TESTING CONVERSION ACTION AVAILABILITY =====");
  const campaignIterator = AdsApp.campaigns()
    .withCondition("Status = ENABLED")
    .get();
  
  if (campaignIterator.hasNext()) {
    const sampleCampaign = campaignIterator.next();
    testConversionActionQuery(sampleCampaign);
  }
  Logger.log("==================================================\n");
}

function testConversionActionQuery(campaign) {
  try {
    // Set dates for last 30 days to make sure we have data
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 30);
    
    const dateRange = {
      start: Utilities.formatDate(startDate, 'UTC', 'yyyyMMdd'),
      end: Utilities.formatDate(endDate, 'UTC', 'yyyyMMdd')
    };
    
    // Query for ALL conversion actions to see what's available
    const query = `
      SELECT 
        campaign.id, 
        metrics.all_conversions,
        segments.conversion_action_name
      FROM campaign 
      WHERE 
        campaign.id = ${campaign.getId()} 
        AND segments.date BETWEEN "${dateRange.start}" AND "${dateRange.end}"
    `;
    
    const report = AdsApp.report(query);
    const rows = report.rows();
    
    Logger.log("Conversion Actions available for campaign " + campaign.getName() + ":");
    
    // List all conversion actions for this campaign
    let totalFound = 0;
    while (rows.hasNext()) {
      const row = rows.next();
      const actionName = row['segments.conversion_action_name'];
      const convCount = parseFloat(row['metrics.all_conversions']) || 0;
      
      if (convCount > 0) {
        Logger.log("  - Action: \"" + actionName + "\" - Count: " + convCount);
        totalFound++;
      }
    }
    
    if (totalFound === 0) {
      Logger.log("  No conversion actions with conversions found in the last 30 days");
    }
    
    return totalFound;
  } catch (e) {
    Logger.log("Error in conversion action test for campaign " + campaign.getName() + ": " + e);
    return 0;
  }
}

function getSpecificConversionMetrics(campaign, dateRange, conversionActionName) {
  try {
    Logger.log(`\n=== Calculating specific conversion metrics for "${campaign.getName()}" ===`);
    Logger.log(`Looking for conversion action: "${conversionActionName}"`);
    Logger.log(`Date range: ${dateRange.start} to ${dateRange.end}`);
    
    // Break the date range into smaller chunks to avoid data issues with long date ranges
    // This will handle 90-day periods by breaking them into monthly chunks
    const startDate = new Date(dateRange.start.substring(0, 4), 
                             parseInt(dateRange.start.substring(4, 6)) - 1, 
                             dateRange.start.substring(6, 8));
    const endDate = new Date(dateRange.end.substring(0, 4), 
                           parseInt(dateRange.end.substring(4, 6)) - 1, 
                           dateRange.end.substring(6, 8));
    
    // Calculate total conversions across all chunks
    let totalConversions = 0;
    let currentStart = new Date(startDate);
    let chunkCount = 0;
    
    // Process in monthly chunks
    while (currentStart <= endDate) {
      chunkCount++;
      // Set end of this chunk (either month end or the overall end date)
      const chunkEnd = new Date(currentStart);
      chunkEnd.setMonth(chunkEnd.getMonth() + 1);
      if (chunkEnd > endDate) {
        chunkEnd.setTime(endDate.getTime());
      }
      
      // Format dates for this chunk
      const chunkDateRange = {
        start: Utilities.formatDate(currentStart, 'UTC', 'yyyyMMdd'),
        end: Utilities.formatDate(chunkEnd, 'UTC', 'yyyyMMdd')
      };
      
      Logger.log(`\nProcessing chunk ${chunkCount}: ${chunkDateRange.start} to ${chunkDateRange.end}`);
      
      // Use a simplified query focusing only on what we need 
      const query = `
        SELECT 
          metrics.all_conversions,
          segments.conversion_action_name
        FROM campaign 
        WHERE 
          campaign.id = ${campaign.getId()} 
          AND segments.date BETWEEN "${chunkDateRange.start}" AND "${chunkDateRange.end}"
      `;
      
      const report = AdsApp.report(query);
      const rows = report.rows();
      
      let chunkConversions = 0;
      // Process each row, carefully tracking the exact conversion action we want
      while (rows.hasNext()) {
        const row = rows.next();
        const actionName = row['segments.conversion_action_name'] || '';
        const convCount = parseFloat(row['metrics.all_conversions']) || 0;
        
        // Only count if it exactly matches our target conversion action name
        if (actionName === conversionActionName) {
          chunkConversions += convCount;
        }
      }
      
      totalConversions += chunkConversions;
      Logger.log(`Chunk ${chunkCount} conversions: ${chunkConversions}`);
      
      // Move to next chunk
      currentStart.setMonth(currentStart.getMonth() + 1);
    }
    
    Logger.log(`\nTotal conversions across ${chunkCount} chunks: ${totalConversions}`);
    Logger.log(`Estimated conversion value: ${totalConversions * CONFIG.CONVERSION_SETTINGS.ESTIMATED_CONVERSION_VALUE}`);
    Logger.log("==================================================\n");
    
    return {
      conversions: totalConversions,
      conversionValue: totalConversions * CONFIG.CONVERSION_SETTINGS.ESTIMATED_CONVERSION_VALUE
    };
  } catch (e) {
    Logger.log("Error getting specific conversion metrics: " + e);
    return { conversions: 0, conversionValue: 0 };
  }
}

function calculateAccountPerformanceScore(campaigns) {
  // Calculate a performance score for the entire account to determine if we should allow a budget increase
  // This considers multiple factors:
  // 1. Overall efficiency improvement (recent vs. long-term)
  // 2. Percentage of campaigns showing improved performance
  // 3. Impression share lost to budget
  
  // Calculate average efficiency ratios
  let totalLongTermEfficiency = 0;
  let totalRecentEfficiency = 0;
  let totalTrendFactor = 0;
  let totalImpressionShareLost = 0;
  let improvingCampaigns = 0;
  
  campaigns.forEach(campaign => {
    totalLongTermEfficiency += campaign.longTermEfficiencyRatio || 0;
    totalRecentEfficiency += campaign.recentEfficiencyRatio || 0;
    totalTrendFactor += campaign.trendFactor || 1.0;
    totalImpressionShareLost += campaign.budgetImpressionShareLost || 0;
    
    // Count campaigns with improving efficiency
    if ((campaign.recentEfficiencyRatio || 0) > (campaign.longTermEfficiencyRatio || 0)) {
      improvingCampaigns++;
    }
  });
  
  const avgLongTermEfficiency = totalLongTermEfficiency / campaigns.length;
  const avgRecentEfficiency = totalRecentEfficiency / campaigns.length;
  const avgTrendFactor = totalTrendFactor / campaigns.length;
  const avgImpressionShareLost = totalImpressionShareLost / campaigns.length;
  const percentImprovingCampaigns = (improvingCampaigns / campaigns.length) * 100;
  
  // Calculate efficiency improvement
  const efficiencyImprovement = avgRecentEfficiency / Math.max(0.1, avgLongTermEfficiency);
  
  // Calculate the overall score - weighting factors that predict good ROI on additional spend
  let score = (efficiencyImprovement * 0.35) + 
             (avgTrendFactor * 0.3) + 
             (Math.min(1, avgImpressionShareLost / 80) * 0.25) + 
             (percentImprovingCampaigns / 100 * 0.1);
             
  // Add a small bonus if more than half the campaigns are improving
  if (percentImprovingCampaigns > 50) {
    score += 0.1;
  }
  
  // Log the calculation components
  Logger.log(`Account performance score calculation:
    - Average efficiency: ${avgLongTermEfficiency.toFixed(2)} → ${avgRecentEfficiency.toFixed(2)} (${efficiencyImprovement.toFixed(2)}x)
    - Average trend factor: ${avgTrendFactor.toFixed(2)}
    - Average impression share lost: ${avgImpressionShareLost.toFixed(2)}%
    - Improving campaigns: ${improvingCampaigns}/${campaigns.length} (${percentImprovingCampaigns.toFixed(1)}%)
    - Overall score: ${score.toFixed(2)}`);
  
  return score;
}

function getRecentCost(campaign, days) {
  try {
    // Calculate date range for the specified number of days
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - days);
    
    const dateRange = {
      start: Utilities.formatDate(startDate, 'UTC', 'yyyyMMdd'),
      end: Utilities.formatDate(endDate, 'UTC', 'yyyyMMdd')
    };
    
    // Get stats for the recent period
    const stats = campaign.getStatsFor(dateRange.start, dateRange.end);
    return stats.getCost();
  } catch (e) {
    Logger.log("Error getting recent cost for campaign " + campaign.getName() + ": " + e);
    return 0;
  }
}

function calculateTrendFactorsForAll(campaignData) {
  Logger.log("Calculating efficiency ratios and trend factors for all campaigns...");
  
  // Calculate total conversions and costs for different time periods
  let totalLongTermConversions = 0;
  let totalLongTermCost = 0;
  let totalRecentConversions = 0;
  let totalRecentCost = 0;
  
  // First, collect totals
  campaignData.campaigns.forEach(campaign => {
    totalLongTermConversions += campaign.conversions;
    totalLongTermCost += campaign.cost;
    totalRecentConversions += campaign.recentConversions;
    
    // Get recent cost if we don't have it yet
    if (!campaign.recentCost) {
      campaign.recentCost = getRecentCost(campaign.campaign, CONFIG.RECENT_PERFORMANCE_PERIOD);
    }
    
    totalRecentCost += campaign.recentCost;
  });
  
  // Then calculate efficiency ratios for each campaign
  campaignData.campaigns.forEach(campaign => {
    // Calculate long-term efficiency ratio (share of conversions / share of cost)
    if (totalLongTermCost > 0 && totalLongTermConversions > 0) {
      const shareOfCost = campaign.cost / totalLongTermCost;
      const shareOfConversions = campaign.conversions / totalLongTermConversions;
      campaign.longTermEfficiencyRatio = shareOfCost > 0 ? shareOfConversions / shareOfCost : 0;
    } else {
      campaign.longTermEfficiencyRatio = 0;
    }
    
    // Calculate recent efficiency ratio
    if (totalRecentCost > 0 && totalRecentConversions > 0) {
      const shareOfRecentCost = campaign.recentCost / totalRecentCost;
      const shareOfRecentConversions = campaign.recentConversions / totalRecentConversions;
      campaign.recentEfficiencyRatio = shareOfRecentCost > 0 ? shareOfRecentConversions / shareOfRecentCost : 0;
    } else {
      campaign.recentEfficiencyRatio = 0;
    }
    
    // Log efficiency ratios
    Logger.log(`Efficiency ratios for '${campaign.name}': 
      - 90-day: ${campaign.longTermEfficiencyRatio.toFixed(2)} (Share of conv: ${(campaign.conversions / Math.max(1, totalLongTermConversions) * 100).toFixed(1)}%, Share of cost: ${(campaign.cost / Math.max(1, totalLongTermCost) * 100).toFixed(1)}%)
      - ${CONFIG.RECENT_PERFORMANCE_PERIOD}-day: ${campaign.recentEfficiencyRatio.toFixed(2)} (Share of conv: ${(campaign.recentConversions / Math.max(1, totalRecentConversions) * 100).toFixed(1)}%, Share of cost: ${(campaign.recentCost / Math.max(1, totalRecentCost) * 100).toFixed(1)}%)`);
  });
  
  // Now calculate trend factors
  campaignData.campaigns.forEach(campaign => {
    campaign.trendFactor = calculateTrendFactor(campaign);
  });
}

/**
 * Simulates day-of-week optimization for previous days to determine if enough 
 * data/confidence would have been available on those days.
 * 
 * @param {Object} campaign - The campaign object
 * @param {number} daysToSimulate - Number of previous days to simulate
 * @param {boolean} includeWeekends - Whether to include weekend days in the simulation
 * @return {Object} - Simulation results for each day
 */
function simulatePreviousDaysOptimization(campaign, daysToSimulate, includeWeekends) {
  const results = {};
  const today = new Date();
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  // Get today's day of week (0-6, Sunday-Saturday)
  const todayDayOfWeek = today.getDay();
  
  // Track which days of the week we've simulated
  const simulatedDays = new Set();
  
  // Simulate each previous day
  for (let i = 1; i <= daysToSimulate; i++) {
    // Create a date object for this previous day
    const simulatedDate = new Date();
    simulatedDate.setDate(today.getDate() - i);
    const dayOfWeek = simulatedDate.getDay(); // 0-6, Sunday-Saturday
    
    // Skip weekends if not included
    if (!includeWeekends && (dayOfWeek === 0 || dayOfWeek === 6)) {
      continue;
    }
    
    // Format the date for logging/display
    const formattedDate = Utilities.formatDate(simulatedDate, 'UTC', 'yyyy-MM-dd');
    
    // Store the original Date.prototype.getDay function
    const originalGetDay = Date.prototype.getDay;
    
    // Override the getDay function to return our simulated day
    Date.prototype.getDay = function() {
      // Check if this is today's date
      if (this.getDate() === today.getDate() && 
          this.getMonth() === today.getMonth() && 
          this.getFullYear() === today.getFullYear()) {
        return dayOfWeek; // Return the simulated day for today
      }
      return originalGetDay.call(this); // Use original for other dates
    };
    
    try {
      // Now run the adaptive day-of-week data collection function
      // It will use our overridden getDay function to simulate a different day
      const simulatedData = getAdaptiveDayOfWeekData(campaign, `simulation_${formattedDate}`);
      
      // Determine if this day would have had enough data/confidence
      const hasReliableData = (simulatedData.adjustment && 
                              simulatedData.adjustment.confidence >= CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MIN_CONFIDENCE_THRESHOLD &&
                              !simulatedData.fallbackApplied);
                              
      const canMakeAdjustment = (simulatedData.adjustment && 
                                simulatedData.adjustment.confidence > 0 &&
                                simulatedData.adjustment.appliedMultiplier !== 1.0);
      
      // Store the results
      results[formattedDate] = {
        date: formattedDate,
        dayName: dayNames[dayOfWeek],
        dayOfWeek: dayOfWeek,
        hasReliableData: hasReliableData,
        canMakeAdjustment: canMakeAdjustment,
        confidenceScore: simulatedData.adjustment ? simulatedData.adjustment.confidence : 0,
        multiplier: simulatedData.adjustment ? simulatedData.adjustment.appliedMultiplier : 1.0,
        lookbackUsed: simulatedData.lookbackUsed,
        usingRelaxedThreshold: simulatedData.usingRelaxedThreshold || false,
        confidenceThreshold: simulatedData.confidenceThreshold || CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MIN_CONFIDENCE_THRESHOLD,
        sampleSize: simulatedData.indices && 
                   simulatedData.indices.dayIndices && 
                   simulatedData.indices.dayIndices[dayOfWeek] ? 
                   simulatedData.indices.dayIndices[dayOfWeek].sampleSize : 0,
        isToday: dayOfWeek === todayDayOfWeek
      };
      
      // Track this day of week as simulated
      simulatedDays.add(dayOfWeek);
      
    } catch (e) {
      // Log any errors but continue with other days
      Logger.log(`Error simulating day-of-week optimization for ${formattedDate}: ${e}`);
    } finally {
      // Restore the original getDay function
      Date.prototype.getDay = originalGetDay;
    }
  }
  
  // Log simulation summary
  Logger.log("\n=== Day-of-Week Simulation Summary ===");
  Logger.log(`Simulated ${Object.keys(results).length} different days`);
  Logger.log(`Covered ${simulatedDays.size} unique days of the week`);
  
  const daysWithReliableData = Object.values(results).filter(r => r.hasReliableData).length;
  const daysWithAdjustments = Object.values(results).filter(r => r.canMakeAdjustment).length;
  
  Logger.log(`Days with reliable data: ${daysWithReliableData}`);
  Logger.log(`Days with significant adjustments: ${daysWithAdjustments}`);
  
  // Group results by day of week for better analysis
  const resultsByDay = {};
  Object.values(results).forEach(result => {
    if (!resultsByDay[result.dayName]) {
      resultsByDay[result.dayName] = [];
    }
    resultsByDay[result.dayName].push(result);
  });
  
  // Log detailed analysis by day of week
  Logger.log("\n=== Detailed Analysis by Day of Week ===");
  Object.entries(resultsByDay).forEach(([dayName, dayResults]) => {
    Logger.log(`\n${dayName}:`);
    Logger.log(`  Total samples: ${dayResults.length}`);
    Logger.log(`  Reliable data: ${dayResults.filter(r => r.hasReliableData).length}`);
    Logger.log(`  Significant adjustments: ${dayResults.filter(r => r.canMakeAdjustment).length}`);
    
    // Calculate average multiplier and confidence
    const avgMultiplier = dayResults.reduce((sum, r) => sum + r.multiplier, 0) / dayResults.length;
    const avgConfidence = dayResults.reduce((sum, r) => sum + r.confidenceScore, 0) / dayResults.length;
    
    Logger.log(`  Average multiplier: ${avgMultiplier.toFixed(2)}`);
    Logger.log(`  Average confidence: ${(avgConfidence * 100).toFixed(1)}%`);
    
    // Log individual samples for this day
    dayResults.forEach(result => {
      Logger.log(`    ${result.date}:`);
      Logger.log(`      - Reliable data: ${result.hasReliableData}`);
      Logger.log(`      - Could make adjustment: ${result.canMakeAdjustment}`);
      Logger.log(`      - Confidence: ${(result.confidenceScore * 100).toFixed(1)}%`);
      Logger.log(`      - Multiplier: ${result.multiplier.toFixed(2)}`);
      Logger.log(`      - Sample size: ${result.sampleSize}`);
      if (result.isToday) {
        Logger.log(`      - This is today's day of week`);
      }
    });
  });
  
  return results;
}

// Find the function that produces the day-of-week performance summary section
// It might be similar to this (look for where "Day-of-week performance analysis for today" is logged)
function logDayOfWeekPerformanceSummary(campaignData) {
  // This is likely where the issue is
  const today = new Date();
  // Fix this to use scriptTimezone
  const dayOfWeek = parseInt(Utilities.formatDate(today, scriptTimezone, "u")) % 7;
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  Logger.log("\n===== DAY-OF-WEEK PERFORMANCE SUMMARY =====");
  Logger.log(`Day-of-week performance analysis for today (${dayNames[dayOfWeek]}):`);
  // ...
}

// Add new function to calculate budget pacing
function calculateBudgetPacing() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth();
  
  // Get the number of days in the current month
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const currentDay = today.getDate();
  
  // Calculate ideal daily budget
  const monthlyBudget = CONFIG.TOTAL_MONTHLY_BUDGET;
  const idealDailyBudget = monthlyBudget / daysInMonth;
  
  // Get actual spend to date using Google Ads API
  const startOfMonth = new Date(year, month, 1);
  const endOfMonth = new Date(year, month + 1, 0);
  
  // Query to get total spend for the month, including shared budgets
  const query = `
    SELECT 
      metrics.cost_micros,
      campaign.id,
      campaign.name
    FROM campaign
    WHERE segments.date BETWEEN "${formatDate(startOfMonth)}" AND "${formatDate(endOfMonth)}"
  `;
  
  const report = AdsApp.report(query);
  const rows = report.rows();
  let totalSpend = 0;
  let sharedBudgetSpend = 0;
  let individualBudgetSpend = 0;
  
  while (rows.hasNext()) {
    const row = rows.next();
    const campaignName = row['campaign.name'];
    const spend = parseFloat(row['metrics.cost_micros']) / 1000000; // Convert micros to actual currency
    
    // Track spend by budget type
    if (campaignName.includes(CONFIG.SHARED_BUDGET_IDENTIFIER)) {
      sharedBudgetSpend += spend;
    } else {
      individualBudgetSpend += spend;
    }
    
    totalSpend += spend;
  }
  
  // Calculate pacing metrics
  const daysElapsed = currentDay;
  const daysRemaining = daysInMonth - currentDay;
  const idealSpendToDate = (monthlyBudget / daysInMonth) * daysElapsed;
  const spendPercentage = (totalSpend / monthlyBudget) * 100;
  
  // Calculate remaining budget and daily targets
  const remainingBudget = monthlyBudget - totalSpend;
  const idealDailyRemaining = remainingBudget / daysRemaining;
  
  // Determine if we're ahead or behind pace
  const paceStatus = totalSpend > idealSpendToDate ? 'ahead' : 'behind';
  const pacePercentage = (totalSpend / idealSpendToDate) * 100;
  
  // Calculate maximum allowed daily budget based on pacing
  let maxDailyBudgetMultiplier = 1.0;
  if (CONFIG.BUDGET_PACING.ALLOW_BUDGET_REDISTRIBUTION) {
    if (paceStatus === 'ahead' && pacePercentage > 105) {
      // If we're significantly ahead, reduce maximum daily budget
      maxDailyBudgetMultiplier = 0.8;
    } else if (paceStatus === 'behind' && pacePercentage < 95) {
      // If we're behind, allow for higher daily budgets
      maxDailyBudgetMultiplier = CONFIG.BUDGET_PACING.MAX_DAILY_BUDGET_MULTIPLIER;
    }
  }
  
  Logger.log("\n===== BUDGET PACING ANALYSIS =====");
  Logger.log(`Current date: ${today.toISOString()}`);
  Logger.log(`Days in month: ${daysInMonth}`);
  Logger.log(`Days elapsed: ${daysElapsed}`);
  Logger.log(`Days remaining: ${daysRemaining}`);
  Logger.log(`Monthly budget: ${monthlyBudget}`);
  Logger.log(`Total spent to date: ${totalSpend.toFixed(2)} (${spendPercentage.toFixed(1)}%)`);
  Logger.log(`  - Individual budgets: ${individualBudgetSpend.toFixed(2)}`);
  Logger.log(`  - Shared budgets: ${sharedBudgetSpend.toFixed(2)}`);
  Logger.log(`Ideal spend to date: ${idealSpendToDate.toFixed(2)}`);
  Logger.log(`Remaining budget: ${remainingBudget.toFixed(2)}`);
  Logger.log(`Ideal daily remaining: ${idealDailyRemaining.toFixed(2)}`);
  Logger.log(`Pace status: ${paceStatus} (${pacePercentage.toFixed(1)}% of ideal)`);
  Logger.log(`Maximum daily budget multiplier: ${maxDailyBudgetMultiplier}`);
  Logger.log("===============================\n");
  
  return {
    totalSpend,
    remainingBudget,
    daysRemaining,
    idealDailyRemaining,
    paceStatus,
    pacePercentage,
    maxDailyBudgetMultiplier,
    sharedBudgetSpend,
    individualBudgetSpend
  };
}

// Add this function before calculateBudgetPacing
/**
 * Formats a date object into YYYYMMDD string format for Google Ads API queries
 * @param {Date} date - The date to format
 * @return {string} - Formatted date string
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

function handleSpecialEvents(campaign, date) {
  const events = [];
  
  // Check for events on this date
  for (const [eventName, event] of Object.entries(CONFIG.SPECIAL_EVENTS)) {
    const eventStart = new Date(event.startDate);
    const eventEnd = new Date(event.endDate);
    
    // Check if date is within event period
    if (date >= eventStart && date <= eventEnd) {
      events.push({
        name: eventName,
        multiplier: event.multiplier,
        confidence: event.confidence,
        description: event.description
      });
    }
    
    // Check impact days
    const impactStart = new Date(eventStart);
    impactStart.setDate(impactStart.getDate() - event.impactDays);
    const impactEnd = new Date(eventEnd);
    impactEnd.setDate(impactEnd.getDate() + event.impactDays);
    
    if (date >= impactStart && date <= impactEnd) {
      // Calculate impact multiplier based on distance from event
      const daysFromEvent = Math.abs(date - eventStart) / (1000 * 60 * 60 * 24);
      const impactMultiplier = 1 + ((event.multiplier - 1) * (1 - (daysFromEvent / event.impactDays)));
      
      events.push({
        name: `${eventName} Impact`,
        multiplier: impactMultiplier,
        confidence: event.confidence * (1 - (daysFromEvent / event.impactDays)),
        description: `${event.description} (${daysFromEvent.toFixed(1)} days from event)`
      });
    }
  }
  
  return events;
}

// Generate a unique cache key for campaign data
function generateCacheKey(campaign, dateRange) {
  try {
    // Create a unique key based on campaign ID, date range, and a hash of the campaign name
    const campaignId = campaign.getId();
    const nameHash = campaign.getName().split('').reduce((acc, char) => {
      return ((acc << 5) - acc) + char.charCodeAt(0) | 0;
    }, 0);
    
    return `${campaignId}_${dateRange.start}_${dateRange.end}_${nameHash}`;
  } catch (e) {
    // Fallback to a simpler key if there's an error
    Logger.log(`Error generating cache key: ${e}`);
    return `${campaign.getId()}_${dateRange.start}_${dateRange.end}`;
  }
}

// Safe campaign data retrieval with fallback
function safeGetCampaignData(campaign, dateRange) {
  try {
    if (!campaign || !dateRange) {
      Logger.log("Invalid campaign or date range provided");
      return null;
    }

    // Get campaign stats with null check
    const stats = campaign.getStatsFor(dateRange.start, dateRange.end);
    if (!stats) {
      Logger.log(`Could not get stats for campaign ${campaign.getName()}`);
      return null;
    }

    // Get budget with null check
    const budget = campaign.getBudget();
    if (!budget) {
      Logger.log(`Could not get budget for campaign ${campaign.getName()}`);
      return null;
    }

    // Get basic metrics with null checks
    const cost = stats.getCost();
    const clicks = stats.getClicks();
    const impressions = stats.getImpressions();

    // Get conversion metrics with null checks
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

    // Calculate key metrics with null checks
    const currentDailyBudget = budget.getAmount();
    const roi = cost > 0 ? conversionValue / cost : 0;
    const cpa = conversions > 0 ? cost / conversions : 0;
    const roas = cost > 0 ? (conversionValue / cost) * 100 : 0;

    // Get performance metrics with null checks
    const impressionShare = getImpressionShare(campaign, dateRange);
    const budgetImpressionShareLost = getImpressionShareLostToBudget(campaign, dateRange);

    // Get day-of-week data if enabled
    let dayOfWeekData = null;
    if (CONFIG.DAY_OF_WEEK.ENABLED) {
      dayOfWeekData = getDayOfWeekData(campaign, dateRange);
    }

    return {
      campaign: campaign,
      name: campaign.getName(),
      currentDailyBudget,
      cost,
      conversions,
      conversionValue,
      clicks,
      impressions,
      roi,
      cpa,
      roas,
      impressionShare,
      budgetImpressionShareLost,
      dayOfWeekData
    };
  } catch (e) {
    Logger.log(`Error in safeGetCampaignData for campaign ${campaign.getName()}: ${e}`);
    return null;
  }
}

function safeCalculateAdjustment(campaign, factors) {
  try {
    if (!campaign) {
      return {
        adjustmentFactor: 1.0,
        confidence: 0.0,
        reason: "Invalid input data"
      };
    }

    // Get adjustment factors directly from campaign object
    const trendFactor = campaign.trendFactor || 1.0;
    const dayOfWeekAdjustment = campaign.dayOfWeekAdjustment || { 
      appliedMultiplier: 1.0, 
      confidence: 0.0 
    };

    // Calculate combined adjustment factor
    let adjustmentFactor = 1.0;
    let confidence = 0.0;
    let reasons = [];

    // Apply trend factor
    if (trendFactor && !isNaN(trendFactor)) {
      adjustmentFactor *= trendFactor;
      reasons.push(`Trend factor: ${trendFactor.toFixed(2)}`);
      confidence += 0.3;
    }

    // Apply day-of-week adjustment
    if (dayOfWeekAdjustment && dayOfWeekAdjustment.appliedMultiplier) {
      adjustmentFactor *= dayOfWeekAdjustment.appliedMultiplier;
      reasons.push(`Day-of-week multiplier: ${dayOfWeekAdjustment.appliedMultiplier.toFixed(2)}`);
      confidence = Math.max(confidence, dayOfWeekAdjustment.confidence || 0.0);
    }

    // Apply impression share lost adjustment when applicable
    if (campaign.budgetImpressionShareLost > 10) {
      const lossMultiplier = 1 + Math.min(0.15, campaign.budgetImpressionShareLost / 100);
      adjustmentFactor *= lossMultiplier;
      reasons.push(`Impression share lost: ${campaign.budgetImpressionShareLost.toFixed(1)}%`);
      confidence += 0.2;
    }

    // Ensure adjustment factor stays within bounds
    adjustmentFactor = Math.max(0.7, Math.min(1.3, adjustmentFactor));
    confidence = Math.min(1.0, confidence);

    return {
      adjustmentFactor: adjustmentFactor,
      confidence: confidence,
      reason: reasons.join(", ")
    };
  } catch (e) {
    Logger.log(`Error calculating adjustment: ${e}`);
    return {
      adjustmentFactor: 1.0,
      confidence: 0.0,
      reason: "Error in calculation"
    };
  }
}

function initializeScript() {
  // Get the account's timezone and make it available throughout the script
  const accountTimezone = AdsApp.currentAccount().getTimeZone();
  scriptTimezone = accountTimezone; // Store in script-level variable
  Logger.log("Account timezone: " + accountTimezone);
  
  // Format today's date in the account timezone
  const today = new Date();
  const todayFormatted = Utilities.formatDate(today, accountTimezone, "EEEE, MMMM d, yyyy");
  Logger.log("Starting Progressive Budget Balancer script...");
  Logger.log("Execution date: " + today.toISOString());
  Logger.log("Today is: " + todayFormatted);
  
  // Log preview mode status
  if (CONFIG.PREVIEW_MODE) {
    Logger.log("\n=== RUNNING IN PREVIEW MODE ===");
    Logger.log("Budget changes will be logged but NOT APPLIED");
    Logger.log("Set CONFIG.PREVIEW_MODE = false to apply changes");
    Logger.log("===================================\n");
  }
}

function getDateRange() {
  // Get date range for analysis (3 months)
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - CONFIG.LOOKBACK_PERIOD);
  
  // Format dates for API queries
  const dateRange = formatDateRange(startDate, endDate);
  Logger.log("Analyzing data from " + dateRange.start + " to " + dateRange.end);
  Logger.log(`Using ${CONFIG.RECENT_PERFORMANCE_PERIOD}-day trending against ${CONFIG.LOOKBACK_PERIOD}-day baseline`);
  
  return dateRange;
}

function getDayOfWeekDateRange() {
  // Get day-of-week analysis date range
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - CONFIG.DAY_OF_WEEK.LOOKBACK_PERIOD);
  const dowDateRange = formatDateRange(startDate, endDate);
  
  // Log day-of-week settings
  if (CONFIG.DAY_OF_WEEK.ENABLED) {
    Logger.log(`Day-of-week optimization is ENABLED (analyzing ${CONFIG.DAY_OF_WEEK.LOOKBACK_PERIOD} days of data)`);
    
    if (CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.ENABLED) {
      Logger.log(`Adaptive lookback is ENABLED (will extend in ${CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.INCREMENT_SIZE}-day increments if needed)`);
      Logger.log(`  - Initial lookback: ${CONFIG.DAY_OF_WEEK.LOOKBACK_PERIOD} days`);
      Logger.log(`  - Maximum lookback: ${CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MAX_TOTAL_LOOKBACK} days`);
      Logger.log(`  - Maximum extensions: ${CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MAX_EXTENSIONS}`);
      Logger.log(`  - Minimum confidence threshold: ${CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.MIN_CONFIDENCE_THRESHOLD}`);
      Logger.log(`  - Minimum data points required: ${CONFIG.DAY_OF_WEEK.MIN_DATA_POINTS}`);
      
      if (CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.PROGRESSIVE_RELAXATION) {
        Logger.log(`  - Progressive relaxation enabled (will try thresholds: ${CONFIG.DAY_OF_WEEK.ADAPTIVE_LOOKBACK.RELAXATION_STEPS.join(', ')})`);
      }
    } else {
      Logger.log("Adaptive lookback is DISABLED (using fixed lookback period)");
    }
  } else {
    Logger.log("Day-of-week optimization is DISABLED");
  }
  
  return dowDateRange;
}

function logConfiguration(dateRange) {
  // Check if we're using specific conversion actions
  if (CONFIG.CONVERSION_SETTINGS.USE_SPECIFIC_CONVERSION_ACTION) {
    Logger.log("Using specific conversion action: " + CONFIG.CONVERSION_SETTINGS.CONVERSION_ACTION_NAME);
  } else {
    Logger.log("Using default conversion actions");
  }
}

function analyzeCampaignPerformance(campaign, dateRange) {
  try {
    // Get day-of-week data first
    const dowDateRange = getDayOfWeekDateRange(dateRange);
    const dayOfWeekData = getDayOfWeekData(campaign, dowDateRange);
    
    // Store day-of-week data directly on campaign object
    campaign.dayOfWeekAdjustment = {
      appliedMultiplier: dayOfWeekData.appliedMultiplier,
      confidence: dayOfWeekData.confidence,
      rawMultiplier: dayOfWeekData.rawMultiplier
    };
    
    // Get recent conversions (21-day)
    const recentConversions = getRecentConversions(campaign, dateRange);
    campaign.recentConversions = recentConversions;
    
    // Get historical conversions (90-day)
    const historicalConversions = getSpecificConversionMetrics(
      campaign,
      dateRange,
      CONFIG.CONVERSION_SETTINGS.CONVERSION_ACTION_NAME
    ).conversions;
    campaign.historicalConversions = historicalConversions;
    
    // Calculate trend factor
    const trendFactor = calculateTrendFactor(recentConversions, historicalConversions);
    campaign.trendFactor = trendFactor;
    
    // Calculate efficiency metrics
    const efficiencyMetrics = {
      recent: recentConversions > 0 ? campaign.currentDailyBudget / recentConversions : 0,
      historical: historicalConversions > 0 ? campaign.currentDailyBudget / historicalConversions : 0
    };
    campaign.efficiencyMetrics = efficiencyMetrics;
    
    // Calculate performance score
    const performanceScore = calculatePerformanceScore(campaign);
    campaign.performanceScore = performanceScore;
    
    return {
      dayOfWeekData,
      recentConversions,
      historicalConversions,
      trendFactor,
      efficiencyMetrics,
      performanceScore
    };
  } catch (e) {
    Logger.log(`Error analyzing campaign ${campaign.name}: ${e}`);
    return null;
  }
}
