import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration, ISPHttpClientOptions, ISPHttpClientBatchOptions, SPHttpClientBatch, ISPHttpClientBatchCreationOptions } from '@microsoft/sp-http';
import { IDocItemProps } from '../Model/IRecentDocsProps';
import DocItem from '../components/DocItem';
import * as $ from 'jquery';
// import * as React from 'react';
// import * as $ from 'jquery'
// import { escape } from '@microsoft/sp-lodash-subset';

export default class SPClientProvider {
    constructor(context: any) {
        this.context = context;
    }

    private context: any;

    public LoadLatestDocs(pageSize: Number): any {
        const reqHeaders = new Headers();
        reqHeaders.append('odata-version', '3.0');

        return this.context.spHttpClient.get(
            `${this.context.pageContext.site.absoluteUrl}/_api/search/query?querytext='path:"${this.context.pageContext.site.absoluteUrl}" (FileExtension:doc OR FileExtension:docx OR FileExtension:xls OR FileExtension:xlsx OR FileExtension:ppt OR FileExtension:pptx OR FileExtension:pdf) (IsDocument:"True" OR contentclass:"STS_ListItem")'&selectproperties='Title,SecondaryFileExtension,Path,LastModifiedTime,ModifiedBy'&RowsPerPage=4&sortlist='LastModifiedTime:descending'`,
            SPHttpClient.configurations.v1,
            {
                headers: reqHeaders
            }
        ).then((response: SPHttpClientResponse): Promise<{ value: IDocItemProps[] }> => {
            return response.json();
        }).then((data: any): any => {
            let results = data.PrimaryQueryResult.RelevantResults.Table.Rows;
            return this.ConvertToArray(results, pageSize);
        });
    }

    public LoadDocIcon(name: string): any {
        return this.context.spHttpClient.get(
            `${this.context.pageContext.site.absoluteUrl}/_api/web/maptoicon(filename='${name}',progid='',size=1)`,
            SPHttpClient.configurations.v1
        ).then((response: SPHttpClientResponse): Promise<{ value: string }> => {
            return response.json();
        }).then((data: { value: string }): string => {
            const iconUrl = this.GetIcon(name);
            if (iconUrl != null) {
                return iconUrl;
            }
            return `${this.context.pageContext.site.absoluteUrl}/_layouts/15/images/${data.value}`;
        });
    }

    private ConvertToArray(items, pageSize): IDocItemProps[] {
        const loadCount = (Math.floor(pageSize / 5) + 1) * 5;

        let array: IDocItemProps[] = [];

        for (let i = 0; i < items.length; i++) {
            let props = this.GetProperties(items[i].Cells);

            let modified = new Date(props.LastModifiedTime);
            let year = modified.getFullYear();
            let month = modified.getMonth() + 1 < 10 ? "0" + (modified.getMonth() + 1) : modified.getMonth() + 1;
            let date = modified.getDate() < 10 ? "0" + modified.getDate() : modified.getDate();

            array.push({
                title: props.Title + "." + props.SecondaryFileExtension,
                modifyDate: month + "/" + date + "/" + year,
                modifyBy: props.ModifiedBy,
                context: this.context,
                url: props.Path
            });

            if (array.length >= loadCount) {
                break;
            }
        }
        return array;
    }

    private GetIcon(name: string) {
        if (name.lastIndexOf('.') > 0) {
            let extension = name.substring(name.lastIndexOf('.')).toLowerCase();
            if (extension == '.xls' || extension == '.xlsx') {
                return String(require('../../../assets/images/xlsx.png'));
            }
            else if (extension == '.doc' || extension == '.docx') {
                return String(require('../../../assets/images/docx.png'));
            }
            else if (extension == '.ppt' || extension == '.pptx') {
                return String(require('../../../assets/images/pptx.png'));
            }
            else {
                return null;
            }
        }
    }

    private GetProperties(item) {
        let title, secondaryFileExtension, path, lastModifiedTime, modifiedBy;
        item.forEach(function (e) {
            if (e.Key === 'Title') {
                title = e.Value;
            } else if (e.Key === 'SecondaryFileExtension') {
                secondaryFileExtension = e.Value;
            } else if (e.Key === 'LastModifiedTime') {
                lastModifiedTime = e.Value;
            } else if (e.Key === 'ModifiedBy') {
                modifiedBy = e.Value;
            } else if (e.Key === 'Path') {
                path = e.Value;
            }
        });
        return {
            Title: title,
            SecondaryFileExtension: secondaryFileExtension,
            LastModifiedTime: lastModifiedTime,
            ModifiedBy: modifiedBy,
            Path: path
        };
    }
}

