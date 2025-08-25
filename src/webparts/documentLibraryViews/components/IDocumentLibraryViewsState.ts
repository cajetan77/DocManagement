import { IViewInfo } from './IViewInfo';

export interface IDocumentLibraryViewsState {
  views: IViewInfo[];
  isLoading: boolean;
  error: string | null;
  expandedTiles: string[];
  publishedLibraryViews: IViewInfo[];
  workingLibraryViews: IViewInfo[];
}

