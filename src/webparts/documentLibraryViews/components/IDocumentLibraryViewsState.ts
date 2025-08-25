import { IViewInfo } from './IViewInfo';

export interface IDocumentLibraryViewsState {
  views: IViewInfo[];
  isLoading: boolean;
  error: string | null;
  expandedTiles: string[];
  publishedLibraryViews: IViewInfo[];
  workingLibraryViews: IViewInfo[];
  loadingExpandedViews: string[]; // Track which tiles are loading expanded views
}

