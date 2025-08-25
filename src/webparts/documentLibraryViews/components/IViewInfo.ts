export interface IViewInfo {
  Id: string;
  Title: string;
  Url: string;
  DefaultView: boolean;
  ViewType: string;
  Hidden: boolean;
  ServerRelativeUrl: string;
  ItemCount: number;
  PersonalView: boolean;
  // New properties for dashboard-style display
  Status?: string;
  StatusColor?: 'black' | 'blue' | 'orange' | 'red' | 'aliceblue';
  IconName?: string;
  ShowViewMore?: boolean;
}
