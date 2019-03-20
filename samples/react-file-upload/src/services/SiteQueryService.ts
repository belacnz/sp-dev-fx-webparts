import * as strings                                             from 'FileUploadWebPartStrings';
import { IDropdownOption, IPersonaProps, ITag }                 from 'office-ui-fabric-react';
import { SPHttpClient, SPHttpClientResponse }                   from '@microsoft/sp-http';
import { isEmpty }                                              from '@microsoft/sp-lodash-subset';
import { IWebPartContext }                                      from '@microsoft/sp-webpart-base';
import { Text, Log }                                            from '@microsoft/sp-core-library';
import { ISiteQueryService }                                 from './ISiteQueryService';
import { SearchService }                                        from './SearchService';

export class SiteQueryService implements ISiteQueryService {
    private context: IWebPartContext;
    private spHttpClient: SPHttpClient;
    private searchService: SearchService;
    private siteUrlOptions: IDropdownOption[];
    private webUrlOptions: IDropdownOption[];


    constructor(context: IWebPartContext, spHttpClient: SPHttpClient) {
        this.context = context;
        this.spHttpClient = spHttpClient;
        this.searchService = new SearchService(this.spHttpClient);
    }
    public getSiteUrlOptions(): Promise<IDropdownOption[]> {   
        // Resolves the already loaded data if available
        if(this.siteUrlOptions) {
            return Promise.resolve(this.siteUrlOptions);
        }
    
        // Otherwise, performs a REST call to get the data
        return new Promise<IDropdownOption[]>((resolve,reject) => {
            let serverUrl = Text.format("{0}//{1}", window.location.protocol, window.location.hostname); 
    
            this.searchService.getSitesStartingWith(serverUrl)
                .then((urls) => {
                    // Adds the current site collection url to the ones returned by the search (in case the current site isn't indexed yet)
                    this.ensureUrl(urls, this.context.pageContext.site.absoluteUrl);
    
                    // Builds the IDropdownOption[] based on the urls
                    let options:IDropdownOption[] = [ { key: "", text: strings.SiteUrlFieldPlaceholder } ];
                    let urlOptions:IDropdownOption[] = urls.sort().map((url) => { 
                        let serverRelativeUrl = !isEmpty(url.replace(serverUrl, '')) ? url.replace(serverUrl, '') : '/';
                        return { key: url, text: serverRelativeUrl };
                    });
                    options = options.concat(urlOptions);
                    this.siteUrlOptions = options;
                    resolve(options);
                })
                .catch((error) => {
                    reject(error);
                }
            );
        });
    }
    public getWebUrlOptions(siteUrl: string): Promise<IDropdownOption[]> {

        // Resolves an empty array if site is null
        if (isEmpty(siteUrl)) {
            return Promise.resolve(new Array<IDropdownOption>());
        }

        // Resolves the already loaded data if available
        if(this.webUrlOptions) {
            return Promise.resolve(this.webUrlOptions);
        }

        // Otherwise, performs a REST call to get the data
        return new Promise<IDropdownOption[]>((resolve,reject) => {

            this.searchService.getWebsFromSite(siteUrl)
                .then((urls) => {
                    // If querying the current site, adds the current site collection url to the ones returned by the search (in case the current web isn't indexed yet)
                    if(siteUrl.toLowerCase().trim() === this.context.pageContext.site.absoluteUrl.toLowerCase().trim()) {
                        this.ensureUrl(urls, this.context.pageContext.web.absoluteUrl);
                    }
                    
                    // Builds the IDropdownOption[] based on the urls
                    let options:IDropdownOption[] = [ { key: "", text: strings.WebUrlFieldPlaceholder } ];
                    let urlOptions:IDropdownOption[] = urls.sort().map((url) => { 
                        let siteRelativeUrl = !isEmpty(url.replace(siteUrl, '')) ? url.replace(siteUrl, '') : '/';
                        return { key: url, text: siteRelativeUrl };
                    });
                    options = options.concat(urlOptions);
                    this.webUrlOptions = options;
                    resolve(options);
                })
                .catch((error) => {
                    reject(error);
                }
            );
        });
    }
    private ensureUrl(urls: string[], urlToEnsure: string) {
        urlToEnsure = urlToEnsure.toLowerCase().trim();
        let urlExist = urls.filter((u) => { return u.toLowerCase().trim() === urlToEnsure; }).length > 0;

        if(!urlExist) {
            urls.push(urlToEnsure);
        }
    }
    public clearCachedWebUrlOptions() {
        this.webUrlOptions = null;
    }
    private getErrorMessage(webUrl: string, error: any): string {
        let errorMessage:string = error.statusText ? error.statusText : error;
        let serverUrl = Text.format("{0}//{1}", window.location.protocol, window.location.hostname);
        let webServerRelativeUrl = webUrl.replace(serverUrl, '');

        if(error.status === 403) {
            errorMessage = Text.format(strings.ErrorWebAccessDenied, webServerRelativeUrl);
        }
        else if(error.status === 404) {
            errorMessage = Text.format(strings.ErrorWebNotFound, webServerRelativeUrl);
        }
        return errorMessage;
    }

}


