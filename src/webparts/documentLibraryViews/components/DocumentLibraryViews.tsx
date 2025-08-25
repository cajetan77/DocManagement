import * as React from 'react';
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Icon,
  Text
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './DocumentLibraryViews.module.scss';
import type { IDocumentLibraryViewsProps } from './IDocumentLibraryViewsProps';
import type { IDocumentLibraryViewsState } from './IDocumentLibraryViewsState';
import type { IViewInfo } from './IViewInfo';

export default class DocumentLibraryViews extends React.Component<IDocumentLibraryViewsProps, IDocumentLibraryViewsState> {
  constructor(props: IDocumentLibraryViewsProps) {
    super(props);
                               this.state = {
         views: [],
         isLoading: true,
         error: null,
         expandedTiles: [],
         publishedLibraryViews: [],
         workingLibraryViews: [],
         loadingExpandedViews: []
       };
  }

  public componentDidMount(): void {
    // console.log('=== Component Mounted ===');
    // console.log('Initial props:', this.props);
    void this.loadViews();
  }

  public componentDidUpdate(prevProps: IDocumentLibraryViewsProps): void {
    // console.log('=== Component Updated ===');
    // console.log('Previous props:', prevProps);
    // console.log('Current props:', this.props);

    if (prevProps.siteUrl !== this.props.siteUrl || prevProps.listTitle !== this.props.listTitle) {
      // console.log('Props changed, reloading views');
      void this.loadViews();
    }
  }

  private async loadViews(): Promise<void> {
    const { siteUrl, listTitle } = this.props;

    // Debug logging to see what the component received
    // console.log('=== Component Props Debug ===');
    // console.log('Component received siteUrl:', siteUrl);
    // console.log('Component received listTitle:', listTitle);
    // console.log('listTitle.trim():', listTitle ? listTitle.trim() : 'undefined');
    // console.log('Is listTitle empty?', !listTitle || listTitle.trim() === '');

    // Only load views if a library title is specified
    if (!listTitle || listTitle.trim() === '') {
      // console.log('No library title specified, setting empty state');
      this.setState({
        isLoading: false,
        error: null,
        views: []
      });
      return;
    }

    if (!siteUrl) {
      this.setState({
        isLoading: false,
        error: 'Please configure the site URL in the web part properties.'
      });
      return;
    }

    this.setState({ isLoading: true, error: null });

    try {
      // First, let's test what library we're actually connecting to
      await this.testLibraryConnection(siteUrl, listTitle);

      const views = await this.getViews(siteUrl, listTitle);
      this.setState({
        views,
        isLoading: false,
        error: null
      });
    } catch (error) {
      this.setState({
        isLoading: false,
        error: `Failed to load views: ${error.message}`
      });
    }
  }

  private async testLibraryConnection(siteUrl: string, listTitle: string): Promise<void> {
    const { context } = this.props;

    // console.log('=== DEBUG: Testing Library Connection ===');

    // Test 1: Get all lists in the site to see what's available
    try {
      const allListsUrl = `${siteUrl}/_api/web/lists?$select=Title,Id,DefaultViewUrl&$filter=Hidden eq false`;
      // console.log('Testing all lists URL:', allListsUrl);

      const allListsResponse = await context.spHttpClient.get(
        allListsUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (allListsResponse.ok) {
        const allListsData = await allListsResponse.json();
        // console.log('All lists in site:', allListsData.value.map((list: any) => ({
        //   Title: list.Title,
        //   Id: list.Id,
        //   DefaultViewUrl: list.DefaultViewUrl
        // })));

        // Check if our target library exists
        const targetLibrary = allListsData.value.find((list: any) =>
          list.Title.toLowerCase() === listTitle.toLowerCase()
        );

        if (targetLibrary) {
          // console.log('✅ Target library found:', targetLibrary.Title);
        } else {
          // console.log('❌ Target library NOT found. Available libraries:');
          // allListsData.value.forEach((list: any) => {
          //   console.log(`  - "${list.Title}"`);
          // });
        }
      }
    } catch (error) {
      // console.error('Error testing library connection:', error);
    }
  }

  private async getViews(siteUrl: string, listTitle: string): Promise<IViewInfo[]> {
    const { context } = this.props;
    //const startTime = Date.now();

    // console.log('=== DEBUG: Starting view processing ===');
    // console.log('Site URL:', siteUrl);
    // console.log('List Title:', listTitle);

    // First, get the list to ensure it exists
    const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')`;
    // console.log('=== DEBUG: List Details ===');
    // console.log('List URL:', listUrl);

    const listResponse = await context.spHttpClient.get(
      listUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }
    );

    if (!listResponse.ok) {
      throw new Error(`List '${listTitle}' not found. Please check the list title.`);
    }

    //const listData = await listResponse.json();
    // console.log('List found:', listData.Title);
    // console.log('List ID:', listData.Id);
    // console.log('List URL:', listData.DefaultViewUrl);

    // Get all views for the list
    const viewsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/views`;
    // console.log('=== DEBUG: Views API Call ===');
    // console.log('Views URL:', viewsUrl);

    const response: SPHttpClientResponse = await context.spHttpClient.get(
      viewsUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }
    );

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    // console.log('=== DEBUG: Raw API Response ===');
    // console.log('API Response Status:', response.status);
    // console.log('API Response Headers:', response.headers);
    // console.log('Full API Response Data:', data);

    const views = data.value || [];

    // console.log('=== DEBUG: All Views Found ===');
    // console.log('Total views found:', views.length);
    // views.forEach((view: any, index: number) => {
    //   console.log(`View ${index + 1}: "${view.Title}" (ID: ${view.Id}, Default: ${view.DefaultView}, Hidden: ${view.Hidden})`);
    // });
    // console.log('Raw views data:', views);

    // Get item count for each view using the view's URL and filter out personal views
    const viewsWithCounts = await Promise.all(
      views.map(async (view: any) => {
        try {
          let itemCount = 0;
          let isPersonalView = false;

          if (view.Id) {
            // Try to get the view's query and use it to count items
            const viewDetailsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/views('${view.Id}')?$select=ViewQuery,PersonalView,DefaultView`;
            // console.log('View details URL:', viewDetailsUrl);
            const viewDetailsResponse = await context.spHttpClient.get(
              viewDetailsUrl,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  'Accept': 'application/json;odata=nometadata',
                  'Content-type': 'application/json;odata=nometadata',
                  'odata-version': ''
                }
              }
            );

            if (viewDetailsResponse.ok) {
              const viewDetails: any = await viewDetailsResponse.json();
              // console.log(`View details for "${view.Title}":`, viewDetails);

              // Check if it's a personal view
              isPersonalView = viewDetails.PersonalView || false;

              // Check if this is a rejected status view that needs special handling
              const title = view.Title.toLowerCase();
              if (title.indexOf('pending') >= 0) {
                // Get count for items with pending moderation status
                try {
                  itemCount = await this.getPendingModerationCount(siteUrl, listTitle);
                  // console.log(`View "${view.Title}" - Pending moderation status count: ${itemCount}`);
                } catch (error) {
                  // console.error(`Error getting pending count for "${view.Title}":`, error);
                  itemCount = 0;
                }
              }
              else if (title.indexOf('rejected') >= 0) {
                // Get count for items with rejected status
                try {
                  itemCount = await this.getRejectedStatusCount(siteUrl, listTitle);
                  // console.log(`View "${view.Title}" - Rejected status count: ${itemCount}`);
                } catch (error) {
                  // console.error(`Error getting rejected count for "${view.Title}":`, error);
                  itemCount = 0;
                }
              }
              else if (title.indexOf('draft') >= 0) {
                // Get count for items with rejected status
                try {
                  itemCount = await this.getDraftStatusCount(siteUrl, listTitle);
                  // console.log(`View "${view.Title}" - Rejected status count: ${itemCount}`);
                } catch (error) {
                  // console.error(`Error getting rejected count for "${view.Title}":`, error);
                  itemCount = 0;
                }
              }
              else if (title.indexOf('assigned') >= 0) {
                // Get count for items with rejected status
                try {
                  itemCount = await this.getMyassignedStatusCount(siteUrl, listTitle); 
                  // console.log(`View "${view.Title}" - Rejected status count: ${itemCount}`);
                } catch (error) {
                  // console.error(`Error getting rejected count for "${view.Title}":`, error);
                  itemCount = 0;
                }
              }
              else {
                // Get accurate item count using RenderListDataAsStream with user context
                try {
                  const renderListDataUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/RenderListDataAsStream`;
                  const renderListDataBody = {
                    parameters: {
                      ViewXml: viewDetails.ViewQuery || '',
                      ViewId: view.Id,
                      RenderOptions: 1, // Include item count
                      OverrideViewXml: false,
                      AddRequiredFields: true,
                      AddRequiredColumns: true
                    }
                  };

                  const renderResponse = await context.spHttpClient.post(
                    renderListDataUrl,
                    SPHttpClient.configurations.v1,
                    {
                      headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=nometadata',
                        'odata-version': ''
                      },
                      body: JSON.stringify(renderListDataBody)
                    }
                  );

                                     if (renderResponse.ok) {
                     const renderData: any = await renderResponse.json();
                     // console.log(`View "${view.Title}" - RenderListDataAsStream response:`, renderData);

                     // Extract items from the response
                     let items: any[] = [];
                     if (renderData.Row && Array.isArray(renderData.Row)) {
                       items = renderData.Row;
                     } else if (renderData.ListData && renderData.ListData.Row) {
                       items = renderData.ListData.Row;
                     }

                     // Filter items based on user permissions
                     const accessibleItems = await this.filterItemsByUserPermissions(items, siteUrl, listTitle);
                     itemCount = accessibleItems.length;

                     // console.log(`View "${view.Title}" - View-specific count (user-filtered): ${itemCount}`);
                   } else {
                     // console.warn(`Failed to get view-specific count for "${view.Title}":`, renderResponse.status, renderResponse.statusText);

                    // Fallback to user-filtered list count
                    try {
                      itemCount = await this.getUserFilteredListCount(siteUrl, listTitle);
                      // console.log(`View "${view.Title}" - Fallback to user-filtered count: ${itemCount}`);
                    } catch (fallbackError) {
                      // console.error(`Error getting fallback count for "${view.Title}":`, fallbackError);
                      itemCount = 0;
                    }
                  }
                } catch (error) {
                  // console.error(`Error getting view-specific count for "${view.Title}":`, error);

                  // Fallback to user-filtered list count
                  try {
                    itemCount = await this.getUserFilteredListCount(siteUrl, listTitle);
                    // console.log(`View "${view.Title}" - Fallback to user-filtered count: ${itemCount}`);
                  } catch (fallbackError) {
                    // console.error(`Error getting fallback count for "${view.Title}":`, fallbackError);
                    itemCount = 0;
                  }
                }
              }
            } else {
              // console.warn(`Failed to get view details for "${view.Title}":`, viewDetailsResponse.status, viewDetailsResponse.statusText);
            }
          }

          return {
            Id: view.Id || '',
            Title: view.Title || 'Untitled View',
            Url: view.Url || '',
            DefaultView: view.DefaultView || false,
            ViewType: view.ViewType || 'HTML',
            Hidden: view.Hidden || false,
            ServerRelativeUrl: view.ServerRelativeUrl || '',
            ItemCount: itemCount,
            PersonalView: isPersonalView
          };
        } catch (error) {
          // console.error('Error processing view:', view, error);
          return {
            Id: view.Id || '',
            Title: view.Title || 'Untitled View',
            Url: view.Url || '',
            DefaultView: view.DefaultView || false,
            ViewType: view.ViewType || 'HTML',
            Hidden: view.Hidden || false,
            ServerRelativeUrl: view.ServerRelativeUrl || '',
            ItemCount: 0,
            PersonalView: false
          };
        }
      })
    );

    // Filter out personal views and unwanted views, then apply categorization
    const filteredViews = viewsWithCounts
      .filter(view => !view.PersonalView)
      .filter(view => {
        const title = view.Title.toLowerCase();
        // Filter out specific views you don't want to show
        const unwantedViews = [
          'merge documents',
          'relink documents',
          'all documents',
          'assetlibtemp'
        ];
        return !unwantedViews.some(unwanted => title.includes(unwanted));
      })
      .map(view => this.categorizeView(view));

    // Get user-specific document count for the library
    const totalDocumentCount = await this.getUserFilteredListCount(siteUrl, listTitle);
    console.log(`User-specific count for "${listTitle}": ${totalDocumentCount}`);
    
    // Get user-specific published document count from Published Document Library
    const publishedDocumentCount = await this.getUserFilteredListCount(siteUrl, 'Published Documents');
    console.log(`User-specific count for "Published Documents": ${publishedDocumentCount}`);

         // Create summary tiles for document counts
     const summaryTiles: IViewInfo[] = [
      
       {
         Id: 'working-document',
         Title: 'Working Document',
         Url: '',
         DefaultView: false,
         ViewType: 'HTML',
         Hidden: false,
         ServerRelativeUrl: '',
        ItemCount: totalDocumentCount, // User-specific count of documents in Working Document Library
         PersonalView: false,
         Status: 'Working Document',
         StatusColor: 'black',
         IconName: 'Folder',
         ShowViewMore: false
       }
     ];

    // Combine summary tiles with filtered views
    const allViews = [...summaryTiles, ...filteredViews];

    // console.log('=== DEBUG: Personal View Filtering ===');
    // console.log('Total views before filtering:', viewsWithCounts.length);
    // console.log('Personal views found:', viewsWithCounts.filter(v => v.PersonalView).length);
    // console.log('Views after filtering:', filteredViews.length);
    // console.log('Filter results:', viewsWithCounts.map(v => ({ title: v.Title, personal: v.PersonalView, itemCount: v.ItemCount })));
    // console.log('Processed views (excluding personal views):', filteredViews);

    // Additional debugging for rejected views
   // const rejectedViews = filteredViews.filter(v => v.Title.toLowerCase().indexOf('rejected') >= 0);
    // console.log('=== DEBUG: Rejected Views ===');
    // console.log('Rejected views found:', rejectedViews.length);
    // rejectedViews.forEach(v => {
    //   console.log(`Rejected view: "${v.Title}" - Count: ${v.ItemCount}, Status: ${v.Status}`);
    // });

    // console.log('=== DEBUG: Summary Tiles ===');
    // console.log('Summary tiles created:', summaryTiles.length);
    // console.log('Total views including summary tiles:', allViews.length);

   // const endTime = Date.now();
    // console.log(`=== DEBUG: View processing completed in ${endTime - startTime}ms ===`);

    return allViews;
  }

  private async getPendingModerationCount(siteUrl: string, listTitle: string): Promise<number> {
    const { context } = this.props;

    try {
      // Method 1: Try using CAML query with specific field filters and user context
      const camlQuery = {
        query: {
          ViewXml: `
             <View>
               <Query>
                 <Where>
                   <Eq>
                     <FieldRef Name="Status" />
                     <Value Type="Text">Pending</Value>
                   </Eq>
                 </Where>
               </Query>
               <ViewFields>
                 <FieldRef Name="Title" />
                 <FieldRef Name="Status" />
                 <FieldRef Name="FileRef" />
               </ViewFields>
               <RowLimit>5000</RowLimit>
             </View>
           `
        }
      };

      const response = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/GetItems`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: JSON.stringify(camlQuery)
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];

        // Filter items based on user permissions
        const accessibleItems = await this.filterItemsByUserPermissions(items, siteUrl, listTitle);

        // console.log(`Pending moderation status count (user-filtered): ${accessibleItems.length}`);
        return accessibleItems.length;
      }
    } catch (error) {
      // console.error('Error with CAML query approach:', error);
    }

    try {
      // Method 2: Try using REST API with $filter and user context
      const filterUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$filter=_ModerationStatus eq 'Pending'&$select=Title,_ModerationStatus,FileRef&$top=5000`;
      const response = await context.spHttpClient.get(
        filterUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];

        // Filter items based on user permissions
        const accessibleItems = await this.filterItemsByUserPermissions(items, siteUrl, listTitle);

        // console.log(`Pending moderation status count (REST, user-filtered): ${accessibleItems.length}`);
        return accessibleItems.length;
      }
    } catch (error) {
      // console.error('Error with REST filter approach:', error);
    }

    try {
      // Method 3: Try alternative field names with user context
      const alternativeFilters = [
        "_ModerationStatus eq 'Pending'",
        "ModerationStatus eq 'Pending'",
        "Approval_x0020_Status eq 'Pending'",
        "ApprovalStatus eq 'Pending'"
      ];

      for (const filter of alternativeFilters) {
        try {
          const filterUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(filter)}')/items?$filter=${encodeURIComponent(filter)}&$select=Title,_ModerationStatus,ModerationStatus,Approval_x0020_Status,ApprovalStatus,FileRef&$top=5000`;
          const response = await context.spHttpClient.get(
            filterUrl,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata',
                'odata-version': ''
              }
            }
          );

          if (response.ok) {
            const data = await response.json();
            const items = data.value || [];

            // Filter items based on user permissions
            const accessibleItems = await this.filterItemsByUserPermissions(items, siteUrl, listTitle);

            // console.log(`Pending moderation status count (${filter}, user-filtered): ${accessibleItems.length}`);
            return accessibleItems.length;
          }
        } catch (filterError) {
          // console.log(`Filter "${filter}" failed:`, filterError);
        }
      }
    } catch (error) {
      // console.error('Error with alternative field names:', error);
    }

    // Fallback: Get all items and filter in JavaScript with user permissions
    try {
      const allItemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=Title,Status,Approval_x0020_Status,ApprovalStatus,_ModerationStatus,ModerationStatus,FileRef&$top=5000`;
      const response = await context.spHttpClient.get(
        allItemsUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];

        // Filter items that have _ModerationStatus = "Pending"
        const pendingItems = items.filter((item: any) => {
          const moderationStatus = item._ModerationStatus || item.ModerationStatus || item.Approval_x0020_Status || item.ApprovalStatus || '';
          return moderationStatus.toLowerCase() === 'pending';
        });

        // Filter items based on user permissions
        const accessibleItems = await this.filterItemsByUserPermissions(pendingItems, siteUrl, listTitle);

        // console.log(`Pending moderation status count (JavaScript filter, user-filtered): ${accessibleItems.length}`);
        return accessibleItems.length;
      }
      else {
        // console.log('Error with JavaScript filtering approach:', response.status, response.statusText);
      }
    } catch (error) {
      // console.error('Error with JavaScript filtering approach:', error);
    }

    // console.warn('Could not determine pending moderation count, returning 0');
    return 0;
  }

  private async getRejectedStatusCount(siteUrl: string, listTitle: string): Promise<number> {
    const { context } = this.props;

    try {
      // Use REST API with $filter to get items where Status column is "Rejected"
      const filterUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$filter=Status eq 'Rejected'&$select=Title,Status,FileRef&$top=5000`;
      const response = await context.spHttpClient.get(
        filterUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];

        // Filter items based on user permissions
        const accessibleItems = await this.filterItemsByUserPermissions(items, siteUrl, listTitle);

        // console.log(`Rejected status count (user-filtered): ${accessibleItems.length}`);
        return accessibleItems.length;
      }
    } catch (error) {
      // console.error('Error getting rejected status count:', error);
    }

    // Fallback: Get all items and filter in JavaScript
    try {
      const allItemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=Title,Status,FileRef&$top=5000`;
      const response = await context.spHttpClient.get(
        allItemsUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];

        // Filter items that have Status = "Rejected"
        const rejectedItems = items.filter((item: any) => {
          const status = item.Status || '';
          return status.toLowerCase() === 'rejected';
        });

        // Filter items based on user permissions
        const accessibleItems = await this.filterItemsByUserPermissions(rejectedItems, siteUrl, listTitle);

        // console.log(`Rejected status count (JavaScript filter, user-filtered): ${accessibleItems.length}`);
        return accessibleItems.length;
      }
    } catch (error) {
      // console.error('Error with JavaScript filtering approach:', error);
    }

    // console.warn('Could not determine rejected status count, returning 0');
    return 0;
  }

  private async getDraftStatusCount(siteUrl: string, listTitle: string): Promise<number> {
    const { context } = this.props;

    try {
      // Use REST API with $filter to get items where Status column is "Rejected"
      const filterUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$filter=Status eq 'Draft'&$select=Title,Status,FileRef&$top=5000`;
      const response = await context.spHttpClient.get(
        filterUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];

        // Filter items based on user permissions
        const accessibleItems = await this.filterItemsByUserPermissions(items, siteUrl, listTitle);

        // console.log(`Rejected status count (user-filtered): ${accessibleItems.length}`);
        return accessibleItems.length;
      }
    } catch (error) {
      // console.error('Error getting rejected status count:', error);
    }

    // Fallback: Get all items and filter in JavaScript
    try {
      const allItemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=Title,Status,FileRef&$top=5000`;
      const response = await context.spHttpClient.get(
        allItemsUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];

        // Filter items that have Status = "Rejected"
        const draftItems = items.filter((item: any) => {
          const status = item.Status || '';
          return status.toLowerCase() === 'draft';
        });

        // Filter items based on user permissions
        const accessibleItems = await this.filterItemsByUserPermissions(draftItems, siteUrl, listTitle);

        // console.log(`Draft status count (JavaScript filter, user-filtered): ${accessibleItems.length}`);
        return accessibleItems.length;
      }
    } catch (error) {
      // console.error('Error with JavaScript filtering approach:', error);
    }

    // console.warn('Could not determine rejected status count, returning 0');
    return 0;
  }

  private async getMyassignedStatusCount(siteUrl: string, listTitle: string): Promise<number> {
    const { context } = this.props;

    try {
      // Get current user ID
      const currentUserUrl = `${siteUrl}/_api/web/currentuser`;
      const userResponse = await context.spHttpClient.get(
        currentUserUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!userResponse.ok) {
        throw new Error('Failed to get current user');
      }

      const userData = await userResponse.json();
      const userId = userData.Id;
      console.log('userId', userId);
      // Use REST API with $filter to get items where AssignedTo column is the current user
      const filterUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$filter=AssignedTo eq '${userId}'&$select=Title,Status,FileRef&$top=5000`;
      console.log('Assigned to me filterUrl', filterUrl);
      const response = await context.spHttpClient.get(
        filterUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];

        // Filter items based on user permissions
        const accessibleItems = await this.filterItemsByUserPermissions(items, siteUrl, listTitle);

        // console.log(`Rejected status count (user-filtered): ${accessibleItems.length}`);
        return accessibleItems.length;
      }
    } catch (error) {
      // console.error('Error getting rejected status count:', error);
    }

    // Fallback: Get all items and filter in JavaScript
    try {
      const allItemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=Title,Status,FileRef&$top=5000`;
      const response = await context.spHttpClient.get(
        allItemsUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];

        // Filter items that have Status = "Rejected"
        const draftItems = items.filter((item: any) => {
          const status = item.Status || '';
          return status.toLowerCase() === 'draft';
        });

        // Filter items based on user permissions
        const accessibleItems = await this.filterItemsByUserPermissions(draftItems, siteUrl, listTitle);

        // console.log(`Draft status count (JavaScript filter, user-filtered): ${accessibleItems.length}`);
        return accessibleItems.length;
      }
    } catch (error) {
      // console.error('Error with JavaScript filtering approach:', error);
    }

    // console.warn('Could not determine rejected status count, returning 0');
    return 0;
  }


  private async filterItemsByUserPermissions(items: any[], siteUrl: string, listTitle: string): Promise<any[]> {
   // const { context } = this.props;

    // If no items, return empty array
    if (!items || items.length === 0) {
      return [];
    }

    // Get current user info
   // const currentUser = context.pageContext.user;
    // console.log('Current user:', currentUser.loginName);
    // console.log(`Filtering ${items.length} items for user permissions`);

    // For performance reasons, we'll assume the user has access to items returned by the API
    // since SharePoint REST API already respects user permissions when called with user context
    // The SPHttpClient automatically includes user authentication, so items returned should be accessible
    
    // Return all items (assuming SharePoint API already filtered by permissions)
    // console.log(`Assuming user has access to all ${items.length} items (SharePoint API permission filtering)`);
    return items;

    /* 
    // Uncomment this section if you need strict permission checking
    // Note: This can be very slow for large lists as it checks each item individually
    const accessibleItems: any[] = [];
    
    // Check permissions for each item (this can be slow for large lists)
    for (const item of items) {
      try {
        const itemUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${item.Id})?$select=Id,Title`;
        const response = await context.spHttpClient.get(
          itemUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }
        );
        
        if (response.ok) {
          accessibleItems.push(item);
        }
      } catch (error) {
        console.log(`User doesn't have access to item ${item.Id}:`, error);
      }
    }
    
    console.log(`User has access to ${accessibleItems.length} out of ${items.length} items`);
    return accessibleItems;
    */
  }

  private async getUserFilteredListCount(siteUrl: string, listTitle: string): Promise<number> {
    const { context } = this.props;

    try {
      // Get all items from the list with user context
      const allItemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=Id,Title,FileRef&$top=5000`;
      const response = await context.spHttpClient.get(
        allItemsUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];

        // Filter items based on user permissions
        const accessibleItems = await this.filterItemsByUserPermissions(items, siteUrl, listTitle);

         console.log(`User-filtered list count for "${listTitle}": ${accessibleItems.length}`);
        return accessibleItems.length;
      }
    } catch (error) {
      // console.error('Error getting user-filtered list count:', error);
    }

    // Fallback: Try to get a basic count if the above fails
    try {
      const basicCountUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/ItemCount`;
      const response = await context.spHttpClient.get(
        basicCountUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const count = data.value || 0;
        // console.log(`Basic count for "${listTitle}": ${count}`);
        return count;
      }
    } catch (fallbackError) {
      // console.error('Error getting basic count:', fallbackError);
    }

    return 0;
  }

          /* private async loadPublishedDocumentLibraryViews(): Promise<void> {
      const { siteUrl, context } = this.props;
      
      try {
        console.log('=== loadPublishedDocumentLibraryViews Debug ===');
        console.log('Site URL:', siteUrl);
        console.log('Loading views from Published Documents Library...');
        
        // Direct API call to get views from Published Documents library
        const viewsUrl = `${siteUrl}/_api/web/lists/getbytitle('Published%20Documents')/views`;
        console.log('Views URL:', viewsUrl);
        
        const response = await context.spHttpClient.get(
          viewsUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }
        );
        
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        const data = await response.json();
        console.log('Raw API Response for Published Documents:', data);
        
        const views = data.value || [];
        console.log('Raw views from Published Documents:', views);
        
        // Convert to IViewInfo format
        const publishedViews: IViewInfo[] = views.map((view: any) => ({
          Id: view.Id || '',
          Title: view.Title || 'Untitled View',
          Url: view.Url || '',
          DefaultView: view.DefaultView || false,
          ViewType: view.ViewType || 'HTML',
          Hidden: view.Hidden || false,
          ServerRelativeUrl: view.ServerRelativeUrl || '',
          ItemCount: 0, // We'll get this separately if needed
          PersonalView: false,
          Status: view.Title,
          StatusColor: 'black',
          IconName: 'Document',
          ShowViewMore: false
        }));
        
        console.log('Converted published views:', publishedViews);
        
        // Filter out unwanted views
        const filteredPublishedViews = publishedViews.filter(view => {
          const title = view.Title.toLowerCase();
          const unwantedViews = [
            'merge documents',
            'relink documents',
            'all documents',
            'assetlibtemp'
          ];
          return !unwantedViews.some(unwanted => title.indexOf(unwanted) >= 0);
        });
        
        console.log('Filtered published views:', filteredPublishedViews);
        this.setState({ publishedLibraryViews: filteredPublishedViews });
        console.log('Published Document Library views loaded:', filteredPublishedViews.length);
      } catch (error) {
        console.error('Error loading Published Document Library views:', error);
        this.setState({ publishedLibraryViews: [] });
      }
    }*/

    private async loadWorkingDocumentLibraryViews(): Promise<void> {
      const { siteUrl, context, listTitle } = this.props;
      
      try {
        console.log('=== loadWorkingDocumentLibraryViews Debug ===');
        console.log('Site URL:', siteUrl);
        console.log('List Title:', listTitle);
        console.log('Loading views from Working Document Library...');
        
        // Direct API call to get views from Working Document library
        const viewsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/views`;
        console.log('Views URL:', viewsUrl);
        
        const response = await context.spHttpClient.get(
          viewsUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }
        );
        
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        const data = await response.json();
        console.log('Raw API Response for Working Documents:', data);
        
        const views = data.value || [];
        console.log('Raw views from Working Documents:', views);
        
        // Get item count for each view using the view's URL and filter out personal views
        const viewsWithCounts = await Promise.all(
          views.map(async (view: any) => {
            try {
              let itemCount = 0;
              let isPersonalView = false;

              if (view.Id) {
                // Try to get the view's query and use it to count items
                const viewDetailsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/views('${view.Id}')?$select=ViewQuery,PersonalView,DefaultView`;
                const viewDetailsResponse = await context.spHttpClient.get(
                  viewDetailsUrl,
                  SPHttpClient.configurations.v1,
                  {
                    headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'Content-type': 'application/json;odata=nometadata',
                      'odata-version': ''
                    }
                  }
                );

                                 if (viewDetailsResponse.ok) {
                   const viewDetails: any = await viewDetailsResponse.json();
                   
                   // Check if it's a personal view
                   isPersonalView = viewDetails.PersonalView || false;

                   // Check if this is a status view that needs special handling
                   const title = view.Title.toLowerCase();
                   if (title.indexOf('pending') >= 0) {
                     // Get count for items with pending moderation status
                     try {
                       itemCount = await this.getPendingModerationCount(siteUrl, listTitle);
                       console.log(`View "${view.Title}" - Pending moderation status count: ${itemCount}`);
                     } catch (error) {
                       console.error(`Error getting pending count for "${view.Title}":`, error);
                       itemCount = 0;
                     }
                   }
                   else if (title.indexOf('rejected') >= 0) {
                     // Get count for items with rejected status
                     try {
                       itemCount = await this.getRejectedStatusCount(siteUrl, listTitle);
                       console.log(`View "${view.Title}" - Rejected status count: ${itemCount}`);
                     } catch (error) {
                       console.error(`Error getting rejected count for "${view.Title}":`, error);
                       itemCount = 0;
                     }
                   }
                   else if (title.indexOf('draft') >= 0) {
                     // Get count for items with draft status
                     try {
                       itemCount = await this.getDraftStatusCount(siteUrl, listTitle);
                       console.log(`View "${view.Title}" - Draft status count: ${itemCount}`);
                     } catch (error) {
                       console.error(`Error getting draft count for "${view.Title}":`, error);
                       itemCount = 0;
                     }
                   }
                   else if (title.indexOf('assigned') >= 0) {
                     // Get count for items assigned to current user
                     try {
                       itemCount = await this.getMyassignedStatusCount(siteUrl, listTitle); 
                       console.log(`View "${view.Title}" - Assigned to me count: ${itemCount}`);
                     } catch (error) {
                       console.error(`Error getting assigned count for "${view.Title}":`, error);
                       itemCount = 0;
                     }
                   }
                   else {
                     // Get accurate item count using RenderListDataAsStream with user context
                     try {
                       const renderListDataUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/RenderListDataAsStream`;
                       const renderListDataBody = {
                         parameters: {
                           ViewXml: viewDetails.ViewQuery || '',
                           ViewId: view.Id,
                           RenderOptions: 1, // Include item count
                           OverrideViewXml: false,
                           AddRequiredFields: true,
                           AddRequiredColumns: true
                         }
                       };

                       const renderResponse = await context.spHttpClient.post(
                         renderListDataUrl,
                         SPHttpClient.configurations.v1,
                         {
                           headers: {
                             'Accept': 'application/json;odata=nometadata',
                             'Content-type': 'application/json;odata=nometadata',
                             'odata-version': ''
                           },
                           body: JSON.stringify(renderListDataBody)
                         }
                       );

                       if (renderResponse.ok) {
                         const renderData: any = await renderResponse.json();

                         // Extract items from the response
                         let items: any[] = [];
                         if (renderData.Row && Array.isArray(renderData.Row)) {
                           items = renderData.Row;
                         } else if (renderData.ListData && renderData.ListData.Row) {
                           items = renderData.ListData.Row;
                         }

                         // Filter items based on user permissions
                         const accessibleItems = await this.filterItemsByUserPermissions(items, siteUrl, listTitle);
                         itemCount = accessibleItems.length;

                         console.log(`View "${view.Title}" - View-specific count (user-filtered): ${itemCount}`);
                       } else {
                         // Fallback to user-filtered list count
                         try {
                           itemCount = await this.getUserFilteredListCount(siteUrl, listTitle);
                           console.log(`View "${view.Title}" - Fallback to user-filtered count: ${itemCount}`);
                         } catch (fallbackError) {
                           console.error(`Error getting fallback count for "${view.Title}":`, fallbackError);
                           itemCount = 0;
                         }
                       }
                     } catch (error) {
                       console.error(`Error getting view-specific count for "${view.Title}":`, error);

                       // Fallback to user-filtered list count
                       try {
                         itemCount = await this.getUserFilteredListCount(siteUrl, listTitle);
                         console.log(`View "${view.Title}" - Fallback to user-filtered count: ${itemCount}`);
                       } catch (fallbackError) {
                         console.error(`Error getting fallback count for "${view.Title}":`, fallbackError);
                         itemCount = 0;
                       }
                     }
                   }
                 }
              }

              return {
                Id: view.Id || '',
                Title: view.Title || 'Untitled View',
                Url: view.Url || '',
                DefaultView: view.DefaultView || false,
                ViewType: view.ViewType || 'HTML',
                Hidden: view.Hidden || false,
                ServerRelativeUrl: view.ServerRelativeUrl || '',
                ItemCount: itemCount,
                PersonalView: isPersonalView
              };
            } catch (error) {
              console.error('Error processing view:', view, error);
              return {
                Id: view.Id || '',
                Title: view.Title || 'Untitled View',
                Url: view.Url || '',
                DefaultView: view.DefaultView || false,
                ViewType: view.ViewType || 'HTML',
                Hidden: view.Hidden || false,
                ServerRelativeUrl: view.ServerRelativeUrl || '',
                ItemCount: 0,
                PersonalView: false
              };
            }
          })
        );

        // Filter out personal views and unwanted views
        const filteredWorkingViews = viewsWithCounts
          .filter(view => !view.PersonalView)
          .filter(view => {
            const title = view.Title.toLowerCase();
            const unwantedViews = [
              'merge documents',
              'relink documents',
              'all documents',
              'assetlibtemp'
            ];
            return !unwantedViews.some(unwanted => title.indexOf(unwanted) >= 0);
          });
        
                 console.log('Filtered working views with counts:', filteredWorkingViews);
         this.setState(prevState => ({
           workingLibraryViews: filteredWorkingViews,
           loadingExpandedViews: prevState.loadingExpandedViews.filter(id => id !== 'working-document')
         }));
         console.log('Working Document Library views loaded:', filteredWorkingViews.length);
      } catch (error) {
        console.error('Error loading Working Document Library views:', error);
        this.setState({ workingLibraryViews: [] });
      }
    }

  /* private async getTotalDocumentCount(siteUrl: string, listTitle: string): Promise<number> {
    const { context } = this.props;

    try {
      // Get total item count from the list
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')?$filter=FSObjType%20eq%200&$top=5000&$count=true`;
      // console.log('List URL Doc Count:', listUrl);
      const response = await context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const totalCount = data.ItemCount || 0;
        // console.log(`Total document count for library "${listTitle}": ${totalCount}`);
        return totalCount;
      }
    } catch (error) {
      // console.error('Error getting total document count:', error);
    }

    // Fallback to user-filtered count if direct count fails
    try {
      const userFilteredCount = await this.getUserFilteredListCount(siteUrl, listTitle);
      // console.log(`Fallback total document count (user-filtered): ${userFilteredCount}`);
      return userFilteredCount;
    } catch (fallbackError) {
      // console.error('Error getting fallback total document count:', fallbackError);
    }

    return 0;
  }*/



  private categorizeView(view: IViewInfo): IViewInfo {
    const title = view.Title.toLowerCase();

    // Define status categories based on view titles
    if (title.indexOf('published') >= 0 || title.indexOf('final') >= 0) {
      return {
        ...view,
        Status: 'Published Documents',
        StatusColor: 'black',
        IconName: 'Document',
        ShowViewMore: false
      };
    } else if (title.indexOf('working') >= 0 || title.indexOf('active') >= 0 || title.indexOf('current') >= 0) {
      return {
        ...view,
        Status: 'Working Documents',
        StatusColor: 'black',
        IconName: 'Folder',
        ShowViewMore: false
      };
    } else if (title.indexOf('draft') >= 0 || title.indexOf('in progress') >= 0) {
      return {
        ...view,
        Status: view.Title,
        StatusColor: 'blue',
        IconName: 'Edit',
        ShowViewMore: false
      };
    } else if (title.indexOf('pending') >= 0) {
      return {
        ...view,
        Status: view.Title,
        StatusColor: 'orange',
        IconName: 'Clock',
        ShowViewMore: false
      };
    } else if (title.indexOf('awaiting owner') >= 0 || title.indexOf('owner action') >= 0) {
      return {
        ...view,
        Status: 'Documents Awaiting Owner Action',
        StatusColor: 'blue',
        IconName: 'Lock',
        ShowViewMore: false
      };
    } else if (title.indexOf('awaiting review') >= 0 || title.indexOf('review') >= 0) {
      return {
        ...view,
        Status: 'Awaiting Review',
        StatusColor: 'orange',
        IconName: 'Clock',
        ShowViewMore: false
      };
    } else if (title.indexOf('awaiting approval') >= 0 || title.indexOf('approval') >= 0) {
      return {
        ...view,
        Status: 'Awaiting Approval',
        StatusColor: 'orange',
        IconName: 'CheckMark',
        ShowViewMore: false
      };
    } else if (title.indexOf('awaiting formatting') >= 0 || title.indexOf('formatting') >= 0) {
      return {
        ...view,
        Status: 'Awaiting Formatting',
        StatusColor: 'orange',
        IconName: 'Brush',
        ShowViewMore: false
      };
    } else if (title.indexOf('quality') >= 0 || title.indexOf('check') >= 0) {
      return {
        ...view,
        Status: 'Awaiting Quality Team Check',
        StatusColor: 'orange',
        IconName: 'Shield',
        ShowViewMore: false
      };
    } else if (title.indexOf('rejected') >= 0 || title.indexOf('rejection') >= 0) {
      return {
        ...view,
        Status: view.Title,
        StatusColor: 'red',
        IconName: 'Cancel',
        ShowViewMore: false
      };
    } else if (title.indexOf('past review') >= 0 || title.indexOf('overdue') >= 0) {
      return {
        ...view,
        Status: 'Documents Past Review Date',
        StatusColor: 'red',
        IconName: 'Clock',
        ShowViewMore: true
      };
    } else if (title.indexOf('assigned') >= 0 ) {
      return {
        ...view,
        Status: 'Assigned to Me',
        StatusColor: 'red',
        IconName: 'Contact',
        ShowViewMore: false
      };
    } else {
      // Default categorization for other views
      return {
        ...view,
        Status: view.Title,
        StatusColor: 'black',
        IconName: 'DocumentLibrary',
        ShowViewMore: false
      };
    }
  }



           

                    private toggleTileExpansion = (tileId: string): void => {
        this.setState(prevState => {
          const newExpandedTiles = [...prevState.expandedTiles];
          const index = newExpandedTiles.indexOf(tileId);
          const isCurrentlyExpanded = index > -1;
          
          if (isCurrentlyExpanded) {
            newExpandedTiles.splice(index, 1);
            // Remove from loading state when collapsing
            const newLoadingExpandedViews = prevState.loadingExpandedViews.filter(id => id !== tileId);
            return { expandedTiles: newExpandedTiles, loadingExpandedViews: newLoadingExpandedViews };
          } else {
            newExpandedTiles.push(tileId);
            
            // If expanding the Working Document tile, fetch its views
            if (tileId === 'working-document') {
              // Add to loading state
              const newLoadingExpandedViews = [...prevState.loadingExpandedViews, tileId];
              void this.loadWorkingDocumentLibraryViews();
              return { expandedTiles: newExpandedTiles, loadingExpandedViews: newLoadingExpandedViews };
            }
          }
          
          return { expandedTiles: newExpandedTiles, loadingExpandedViews: prevState.loadingExpandedViews };
         });
       };

           private renderExpandedViews = (tileId: string): JSX.Element => {
        const { views, publishedLibraryViews, workingLibraryViews, loadingExpandedViews } = this.state;
        console.log('Regular Views:', views);
      
        // Check if this tile is currently loading expanded views
        const isLoading = loadingExpandedViews.indexOf(tileId) > -1;
      
        // Filter views based on the tile type
        let filteredViews: IViewInfo[] = [];
        let libraryName: string = '';
        
        if (tileId === 'published-documents') {
          // Show views from the Published Document Library
          filteredViews = publishedLibraryViews || [];
          libraryName = 'Published Document Library';
        } else if (tileId === 'working-document') {
          // Show views from the Working Document Library
          filteredViews = workingLibraryViews || [];
          libraryName = 'Working Document Library';
        }
      
        // Show loading spinner if views are being loaded
        if (isLoading) {
          return (
            <div className={styles.expandedViewsContainer}>
              <div className={styles.expandedViewsTitle}>
                <div className={styles.libraryName}>{libraryName} Views:</div>
                <div className={styles.libraryIcon}>
                  <Icon iconName="DocumentLibrary" />
                </div>
              </div>
              <div className={styles.loading}>
                <Spinner size={SpinnerSize.small} label="Loading views..." />
              </div>
            </div>
          );
        }
      
        if (filteredViews.length === 0) {
          return (
            <div className={styles.expandedViewsContainer}>
              <div className={styles.expandedViewsTitle}>
                <div className={styles.libraryName}>{libraryName} Views:</div>
                <div className={styles.libraryIcon}>
                  <Icon iconName="DocumentLibrary" />
                </div>
              </div>
              <div className={styles.emptyState}>
                <Text variant="medium">No views available for this library.</Text>
              </div>
            </div>
          );
        }
     
       return (
         <div className={styles.expandedViewsContainer}>
           <div className={styles.expandedViewsTitle}>
             <div className={styles.libraryName}>{libraryName} Views:</div>
             <div className={styles.libraryIcon}>
               <Icon iconName="DocumentLibrary" />
             </div>
           </div>
           <div className={styles.expandedViewsList}>
             {filteredViews.map((view) => (
               <div
                 key={view.Id}
                 className={styles.expandedViewItem}
                 onClick={() => this.handleExpandedViewClick(view)}
               >
                 <div className={styles.expandedViewTitle}>{view.Title}</div>
                 <div className={styles.expandedViewCount}>{view.ItemCount || 0}</div>
               </div>
             ))}
           </div>
         </div>
       );
     };

     private handleExpandedViewClick = (view: IViewInfo): void => {
       const { siteUrl } = this.props;
       
       // Prefer Url, fallback to ServerRelativeUrl
       let viewUrl: string = view.Url || view.ServerRelativeUrl || '';
       
       // Make relative URLs absolute
       if (viewUrl && viewUrl.indexOf('http') !== 0) {
         const relativePath = viewUrl.indexOf('/') === 0 ? viewUrl.substring(1) : viewUrl;

         const parsed = new URL(siteUrl);
         const rootSite = `${parsed.protocol}//${parsed.host}`;
         viewUrl = `${rootSite}/${relativePath}`;
       }

       if (viewUrl) {
         window.open(viewUrl, '_blank');
       }
     };
 
     private handleViewClick = (view: IViewInfo): void => {
    const { siteUrl, listTitle } = this.props;
     console.log('List Title:', listTitle);
    
          // Handle summary tiles (total documents and working documents)
      if (view.Id === 'published-documents') {
        // Open Published Documents Library
        const publishedLibraryUrl = `${siteUrl}/Published%20Documents`;
        // console.log('Opening Published Documents Library:', publishedLibraryUrl);
        window.open(publishedLibraryUrl, '_blank');
        return;
      } else if (view.Id === 'working-document') {
        // For Working Document, do nothing on tile click - only toggle on arrow click
        const workingLibraryUrl = `${siteUrl}/Working%20Document`;  
        window.open(workingLibraryUrl, '_blank');
        return;
      }

     // Prefer Url, fallback to ServerRelativeUrl
     let viewUrl: string = view.Url || view.ServerRelativeUrl || '';
     // console.log('View URL:', viewUrl);
     // Make relative URLs absolute
     if (viewUrl && viewUrl.indexOf('http') !== 0) {
       const relativePath = viewUrl.indexOf('/') === 0 ? viewUrl.substring(1) : viewUrl;

       const parsed = new URL(siteUrl);
       const rootSite = `${parsed.protocol}//${parsed.host}`;
       // console.log(rootSite);
       viewUrl = `${rootSite}/${relativePath}`;
     }

     if (viewUrl) {
       //console.log('View URL:', viewUrl);
       window.open(viewUrl, '_blank');
     }
   }





   /*private getViewTypeDisplayName(viewType: string): string {
     const viewTypeMap: { [key: string]: string } = {
       'HTML': 'Standard',
       'GRID': 'Grid',
       'CALENDAR': 'Calendar',
       'GANTT': 'Gantt',
       'RECURRENCE': 'Recurrence'
     };
     
     return viewTypeMap[viewType] || viewType;
   }*/

     private renderViewTile = (view: IViewInfo): JSX.Element => {
     // console.log('Rendering view tile:', view);
     const { expandedTiles } = this.state;

     const getStatusColor = (color: string) => {
       switch (color) {
         case 'black': return styles.countBlack;
         case 'blue': return styles.countBlue;
         case 'orange': return styles.countOrange;
         case 'red': return styles.countRed;
         default: return styles.countBlack;
       }
     };

     const isExpanded = expandedTiles.indexOf(view.Id) > -1;
     const isSummaryTile = view.Id === 'published-documents' || view.Id === 'working-document';

         return (
             <div
         key={view.Id}
         className={`${styles.viewTile} ${view.DefaultView ? styles.defaultView : ''} ${(isExpanded && isSummaryTile) ? styles.expandedTile : ''} ${(isExpanded && isSummaryTile) ? styles.hasViewMore : ''}`}
         onClick={() => this.handleViewClick(view)}
                  // title={`Click to ${(view.Id === 'published-documents' || view.Id === 'working-document') ? 'use toggle button' : 'open'} ${view.Status || view.Title} ${(view.Id === 'published-documents' || view.Id === 'working-document') ? 'details' : 'library'}`}
               >
         <div className={styles.tileIcon}>
           <Icon iconName={view.IconName || 'DocumentLibrary'} />
         </div>
         <div className={styles.tileHeader}>
           <div className={`${styles.tileTitle} ${getStatusColor(view.StatusColor || 'black')}`}>{view.Status || view.Title}</div>
         </div>

                          <div className={`${styles.tileCount} ${getStatusColor(view.StatusColor || 'black')}`}>
           {view.ItemCount || 0}
         </div>

                                                                                                                                                               {/* View More button for Published Documents and Working Document */}
                           {(view.Id === 'published-documents' || view.Id === 'working-document') && (
                <div 
                  className={styles.toggleButton}
                  onClick={(e) => {
                    e.stopPropagation(); // Prevent tile click
                    this.toggleTileExpansion(view.Id);
                  }}
                >
                  <span className={styles.toggleText}>
                    {isExpanded ? 'View Less' : 'View More'}
                  </span>
                  <Icon iconName={isExpanded ? 'ChevronUp' : 'ChevronDown'} />
                </div>
              )}

          {view.ShowViewMore && !isSummaryTile && (
            <div className={styles.viewMoreButton}>
              <button
                className={styles.viewMoreBtn}
                onClick={(e) => {
                  e.stopPropagation();
                  this.handleViewClick(view);
                }}
              >
                View more
              </button>
            </div>
          )}

          {/* Expanded views list for Published Documents and Working Document */}
          {(view.Id === 'published-documents' || view.Id === 'working-document') && isExpanded && (
            this.renderExpandedViews(view.Id)
          )}
      </div>
    );
  }



   public render(): React.ReactElement<IDocumentLibraryViewsProps> {

     const { views, isLoading, error } = this.state;

     return (
       <section className={styles.documentLibraryViews}>
         <div className={styles.container}>
           {/*<div className={styles.title}>
         
             {description && (
               <Text variant="medium" style={{ display: 'block', marginTop: '8px', fontWeight: 'normal' }}>
                 {description}
               </Text>
             )}
           </div>*/}


           {isLoading && (
             <div className={styles.loading}>
               <Spinner size={SpinnerSize.large} label="Loading views..." />
             </div>
           )}

           {error && (
             <div className={styles.error}>
               <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                 {error}
               </MessageBar>
             </div>
           )}

           {!isLoading && !error && views.length === 0 && (
             <div className={styles.emptyState}>
               <div className={styles.emptyIcon}>
                 <Icon iconName="DocumentLibrary" />
               </div>
               <Text variant="large">
                 {!this.props.listTitle || this.props.listTitle.trim() === ''
                   ? 'No Document Library Selected'
                   : 'No views found'}
               </Text>
               <Text variant="medium" style={{ marginTop: '8px' }}>
                 {!this.props.listTitle || this.props.listTitle.trim() === ''
                   ? 'Please configure a document library title in the web part properties to display its views.'
                   : `The document library "${this.props.listTitle}" doesn't have any views configured.`}
               </Text>
             </div>
           )}

           {!isLoading && !error && views.length > 0 && (
             <div className={styles.viewsGrid}>
               {views.map(this.renderViewTile)}
             </div>
           )}

           
         </div>
       </section>
     );
   }
}
