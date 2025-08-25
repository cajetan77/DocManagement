# SPFx Web Part Debugging Guide

## 1. Browser Developer Tools Debugging

### Setup:
1. Run the development server:
   ```bash
   gulp serve
   ```

2. Open your SharePoint site and add the web part to a page

3. Open browser developer tools (F12)

### Debugging Steps:

#### Console Logging
The web part already has console.log statements. Check the browser console for:
- Raw views data
- View details for each view
- Processed views (excluding personal views)
- Any errors during API calls

#### Breakpoints
1. In browser dev tools, go to Sources tab
2. Find your web part files under the workbench or your SharePoint site
3. Set breakpoints in the `getViews` method, especially around:
   - Line 108: Where personal views are being processed
   - Line 125: Where `isPersonalView` is set
   - Line 200: Where views are filtered

#### Network Tab
Monitor API calls to see:
- List retrieval requests
- View details requests
- Item count requests

## 2. Enhanced Debugging with Additional Logging

Add these debug statements to your code:

```typescript
// In getViews method, add more detailed logging:
console.log('=== DEBUG: Starting view processing ===');
console.log('Site URL:', siteUrl);
console.log('List Title:', listTitle);

// After filtering personal views:
console.log('=== DEBUG: Personal View Filtering ===');
console.log('Total views before filtering:', viewsWithCounts.length);
console.log('Personal views found:', viewsWithCounts.filter(v => v.PersonalView).length);
console.log('Views after filtering:', filteredViews.length);
```

## 3. SharePoint Workbench Debugging

1. Access the workbench: `https://your-site.sharepoint.com/_layouts/15/workbench.aspx`
2. Add your web part to the workbench
3. Use browser dev tools to debug in isolation

## 4. Visual Studio Code Debugging

### Setup VS Code debugging:

1. Create `.vscode/launch.json`:
```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "name": "Launch Chrome",
      "type": "chrome",
      "request": "launch",
      "url": "https://your-site.sharepoint.com/_layouts/15/workbench.aspx",
      "webRoot": "${workspaceFolder}/src",
      "sourceMaps": true,
      "sourceMapPathOverrides": {
        "webpack:///./src/*": "${webRoot}/*"
      }
    }
  ]
}
```

2. Set breakpoints in VS Code
3. Press F5 to start debugging

## 5. Common Debugging Scenarios

### Personal Views Not Being Filtered:
- Check if `viewDetails.PersonalView` is returning the expected value
- Verify the API call to get view details is successful
- Check if the filter condition `!view.PersonalView` is working

### API Errors:
- Check network tab for failed requests
- Verify site URL and list title are correct
- Check permissions for the current user

### Performance Issues:
- Monitor the number of API calls being made
- Check if item count requests are taking too long
- Consider implementing caching for view details

## 6. Testing Different Scenarios

### Test Cases:
1. **Document library with no personal views**
2. **Document library with personal views**
3. **Document library with mixed view types**
4. **Empty document library**
5. **Document library with hidden views**

### Manual Testing Steps:
1. Create a personal view in SharePoint
2. Add the web part to a page
3. Check if the personal view appears in the list
4. Verify it gets filtered out

## 7. Troubleshooting Tips

### If personal views are still showing:
1. Check the browser console for any errors
2. Verify the `PersonalView` property is being set correctly
3. Add a temporary console.log to see the filter results:
   ```typescript
   console.log('Filter results:', viewsWithCounts.map(v => ({ title: v.Title, personal: v.PersonalView })));
   ```

### If the web part isn't loading:
1. Check if `gulp serve` is running
2. Verify the workbench URL is correct
3. Check browser console for JavaScript errors

### If API calls are failing:
1. Check network tab for HTTP errors
2. Verify the site URL format
3. Check if the list title matches exactly (case-sensitive)

## 8. Performance Monitoring

Add performance logging:
```typescript
const startTime = Date.now();
// ... your code ...
const endTime = Date.now();
console.log(`View processing took ${endTime - startTime}ms`);
```

## 9. Production Debugging

For production issues:
1. Use browser dev tools on the live site
2. Check SharePoint logs (if you have access)
3. Use the browser's network tab to see actual API responses
4. Add temporary logging and deploy to test environment

## 10. Useful Browser Extensions

- **React Developer Tools**: For React component debugging
- **Redux DevTools**: If using Redux
- **SharePoint Framework Dev Tools**: For SPFx-specific debugging
