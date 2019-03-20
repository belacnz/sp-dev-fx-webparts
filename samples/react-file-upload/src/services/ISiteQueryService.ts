import { IDropdownOption, IPersonaProps, ITag } from 'office-ui-fabric-react';

export interface ISiteQueryService {
    getSiteUrlOptions: () => Promise<IDropdownOption[]>;
    getWebUrlOptions: (siteUrl: string) => Promise<IDropdownOption[]>;
    clearCachedWebUrlOptions: () => void;
}