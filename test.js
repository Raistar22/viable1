function testGeminiApiKeyRetrieval() {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (apiKey) {
      Logger.log('GEMINI_API_KEY was retrieved successfully: ' + apiKey.substring(0, 5) + '...'); // Log first 5 chars for safety
    } else {
      Logger.log('GEMINI_API_KEY is NULL or undefined.');
    }
  } catch (e) {
    Logger.log('An error occurred during key retrieval: ' + e.message);
  }
}

/**
 * Test Dashboard System
 */
function testDashboardSystem() {
  console.log('=== Testing Dashboard System ===');
  
  try {
    // Test 1: Check if dashboard functions exist
    console.log('1. Testing function availability...');
    console.log('openDashboard:', typeof openDashboard);
    console.log('openAddClientPopup:', typeof openAddClientPopup);
    console.log('getDashboardData:', typeof getDashboardData);
    console.log('getAllClients:', typeof getAllClients);
    
    // Test 2: Test dashboard data retrieval
    console.log('2. Testing dashboard data retrieval...');
    const dashboardData = getDashboardData();
    console.log('Dashboard data:', dashboardData);
    
    // Test 3: Test client retrieval
    console.log('3. Testing client retrieval...');
    const clients = getAllClients();
    console.log('Clients count:', clients.length);
    console.log('Clients:', clients);
    
    // Test 4: Test metrics
    console.log('4. Testing metrics...');
    const metrics = getDashboardMetrics();
    console.log('Metrics:', metrics);
    
    // Test 5: Test system diagnostics
    console.log('5. Testing system diagnostics...');
    const diagnostics = runSystemDiagnostics();
    console.log('Diagnostics:', diagnostics);
    
    console.log('=== Dashboard System Test Complete ===');
    
    return {
      success: true,
      message: 'Dashboard system test completed successfully',
      results: {
        dashboardData,
        clients: clients.length,
        metrics,
        diagnostics
      }
    };
    
  } catch (error) {
    console.error('Error testing dashboard system:', error);
    return {
      success: false,
      message: 'Error testing dashboard system: ' + error.message,
      error: error
    };
  }
}

/**
 * Test opening the dashboard
 */
function testOpenDashboard() {
  try {
    console.log('Testing dashboard opening...');
    const result = openDashboard();
    console.log('Dashboard opened successfully');
    return { success: true, message: 'Dashboard opened successfully' };
  } catch (error) {
    console.error('Error opening dashboard:', error);
    return { success: false, message: 'Error opening dashboard: ' + error.message };
  }
}

/**
 * Test opening the add client popup
 */
function testOpenAddClientPopup() {
  try {
    console.log('Testing add client popup opening...');
    const result = openAddClientPopup();
    console.log('Add client popup result:', result);
    return result;
  } catch (error) {
    console.error('Error opening add client popup:', error);
    return { success: false, message: 'Error opening popup: ' + error.message };
  }
}