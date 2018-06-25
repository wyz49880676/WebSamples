//import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@microsoft/sp-http';

export interface IRecentDocsWebPartProps {
  title: string
  listUrl: string
  listTitle: string
  siteUrl: string
}
export interface IRecentDocsProps {
  context: any;
  listUrl: string;
  listTitle: string;
  siteUrl: string;
}
export interface IRecentDocsState {
  items: any[];
  parent: string;
  isLoaded: boolean;
}
export interface IDocListProps {
  context: any;
  direct: Direction;
  pageCount: Number;
}
export interface IDocItemProps {
  context: any;
  title: string;
  modifyDate: string;
  modifyBy: string;
  url: string;
  direct?: Direction;
}
export interface IDocListState {
  items: any[];
  isLoaded: boolean;
  message: any;
}
export interface IDocItemState {
  icon: string;
  isLoaded: boolean;
}
export const enum Direction {
  vertical,
  horizonal
}