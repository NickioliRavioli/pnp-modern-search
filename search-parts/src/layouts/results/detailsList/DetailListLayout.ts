import * as React from "react";
import { BaseLayout, IDataContext } from "@pnp/modern-search-extensibility";
import * as strings from 'CommonStrings';
import { IComboBoxOption } from 'office-ui-fabric-react';
import { IDetailsListColumnConfiguration } from '../../../components/DetailsListComponent';
import { 
    IPropertyPaneField, 
    PropertyPaneToggle, 
    PropertyPaneDropdown, 
    PropertyPaneHorizontalRule, 
    PropertyPaneButton, 
    PropertyPaneButtonType, 
    PropertyPaneDropdownOptionType,
    IPropertyPaneDropdownOption, 
} from '@microsoft/sp-property-pane';
import { TemplateValueFieldEditor, ITemplateValueFieldEditorProps } from '../../../controls/TemplateValueFieldEditor/TemplateValueFieldEditor';
import { AsyncCombo } from "../../../controls/PropertyPaneAsyncCombo/components/AsyncCombo";
import { IAsyncComboProps } from "../../../controls/PropertyPaneAsyncCombo/components/IAsyncComboProps";

import { PropertyPaneSlider, PropertyPaneTextField} from '@microsoft/sp-property-pane';
import { IComponentFieldsConfiguration } from '../../../models/common/IComponentFieldsConfiguration';
import { Icon, IIconProps} from 'office-ui-fabric-react';

/**
 * Details List Builtin Layout
 */
export interface IDetailsListLayoutProperties {


    /**
     * The template view to use (locale ID)
     */
    SelectedView: string;

    /**
     * The field used to defined the maxWidth controlling the switch between DetailList and Card views
     */
    responsiveWidth: string;

    // !DetailList Properties
    /**
     * The details list column configuration
     */
    detailsListColumns: IDetailsListColumnConfiguration[];

    /**
     * The field to use for file extension 
     */
    fieldIconExtension: string;

    /**
     * If we should group items by property
     */
    enableGrouping: boolean;

    /**
     * The field used to group items in the list
     */
    groupByField: string;

    /**
     * If groups should collapsed by default
     */
    groupsCollapsed: boolean;

    /**
     * Shows or hide the file icon in the first column
     */
    showFileIcon: boolean;

    /**
     * Show the details list as compact
     */
    isCompact: boolean;

    // !Simple List properties

    /**
     * Show or hide the file icon
     */
    showFileIconList: boolean;

    /**
     * Show or hide the item thumbnail
     */
    showItemThumbnail: boolean;


    // !CARD properties
    /**
     * The document card fields configuration
     */
    documentCardFields: IComponentFieldsConfiguration[];

    /**
     * Indicates of the tile should enable the preview
     */
    enablePreview: boolean;

    /**
     * The prefered number of cards per row
     */
    preferedCardNumberPerRow: number;

    /**
     * The card size in %
     */
    columnSizePercentage: number;

    /**
     * Shows or hide the file icon in the first column
     */
    showFileIconCard: boolean;

    /**
     * Show the details list as compact
     */
    isCompactCard: boolean;


}

export class DetailsListLayout extends BaseLayout<IDetailsListLayoutProperties> {

    /**
     * Dynamically loaded components for property pane
     */
    private _propertyFieldCollectionData: any = null;
    private _propertyPaneWebPartInformation: any = null;
    private _customCollectionFieldType: any = null;

    private _propertyFieldToogleWithCallout: any = null;
    private _propertyFieldCalloutTriggers: any = null;

    public async onInit(): Promise<void> {

        this.properties.SelectedView = this.properties.SelectedView ? this.properties.SelectedView : 'Simple List';

        this.properties.responsiveWidth = this.properties.responsiveWidth != null ? this.properties.responsiveWidth : "640";

        // Setup default values
        this.properties.detailsListColumns = this.properties.detailsListColumns ? this.properties.detailsListColumns :
            [
                {
                    name: 'Title',
                    value: '<a href="{{slot item @root.slots.PreviewUrl}}" target="_blank" style="color: {{@root.theme.semanticColors.link}}">\n\t{{slot item @root.slots.Title}}\n</a>',
                    useHandlebarsExpr: true,
                    minWidth: '80',
                    maxWidth: '300',
                    enableSorting: false,
                    isMultiline: false,
                    isResizable: true
                },
                {
                    name: 'Created',
                    value: "{{getDate (slot item @root.slots.Date) 'LL'}}",
                    useHandlebarsExpr: true,
                    minWidth: '80',
                    maxWidth: '120',
                    enableSorting: false,
                    isMultiline: false,
                    isResizable: false
                },
                {
                    name: 'Summary',
                    value: "{{getSummary (slot item @root.slots.Summary)}}",
                    useHandlebarsExpr: true,
                    minWidth: '80',
                    maxWidth: '300',
                    enableSorting: false,
                    isMultiline: true,
                    isResizable: false
                }
            ] as IDetailsListColumnConfiguration[];

        this.properties.isCompact = this.properties.isCompact !== null && this.properties.isCompact !== undefined ? this.properties.isCompact : false;
        this.properties.showFileIcon = this.properties.showFileIcon !== null && this.properties.showFileIcon !== undefined ? this.properties.showFileIcon : true;
        this.properties.fieldIconExtension = this.properties.fieldIconExtension ? this.properties.fieldIconExtension : 'FileType';
        this.properties.enableGrouping = this.properties.enableGrouping !== null && this.properties.enableGrouping !== undefined ? this.properties.enableGrouping : false;
        this.properties.groupByField = this.properties.groupByField ? this.properties.groupByField : '';
        this.properties.groupsCollapsed = this.properties.groupsCollapsed !== null && this.properties.groupsCollapsed !== undefined ? this.properties.groupsCollapsed : true;

        const { PropertyFieldCollectionData, CustomCollectionFieldType } = await import(
            /* webpackChunkName: 'pnp-modern-search-results-detailslist-layout' */
            '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
        );
        this._propertyFieldCollectionData = PropertyFieldCollectionData;

        if (this.properties.enableGrouping) {
            const { PropertyPaneWebPartInformation } = await import(
                /* webpackChunkName: 'pnp-modern-search-property-pane' */
                '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation'
            );
            this._propertyPaneWebPartInformation = PropertyPaneWebPartInformation;
        }

        this._customCollectionFieldType = CustomCollectionFieldType;


        this.properties.showFileIconList = this.properties.showFileIconList !== null && this.properties.showFileIconList !== undefined ?  this.properties.showFileIconList: true;
        this.properties.showItemThumbnail = this.properties.showItemThumbnail !== null && this.properties.showItemThumbnail !== undefined ?  this.properties.showItemThumbnail: true;
    


        this.properties.documentCardFields = this.properties.documentCardFields ? this.properties.documentCardFields :
                                                [
                                                    { name: strings.Layouts.Cards.Fields.Title, field: 'title', value: '{{slot item @root.slots.Title}}', useHandlebarsExpr: true, supportHtml: false },
                                                    { name: strings.Layouts.Cards.Fields.Location, field: 'location', value: `<a style="color:{{@root.theme.palette.themePrimary}};font-weight:600;font-family:'{{@root.theme.fonts.small.fontFamily}}'" href="{{SPSiteUrl}}">{{SiteTitle}}</a>`, useHandlebarsExpr: true, supportHtml: true },
                                                    { name: strings.Layouts.Cards.Fields.Tags, field: 'tags', value: `<style>\n\t.tags {\n\t\tdisplay: flex;\n\t\talign-items: center;\n\t }\n\t.tags i { \n\t\tmargin-right: 5px; \n\t}\n\t.tags div {\n\t\tdisplay: flex;\n\t\tflex-wrap: wrap; \n\t\tjustify-content: flex-end; \n\t}\n\t.tags div span {\n\t\ttext-decoration: underline; \n\t\tmargin-right: auto; \n\t}\n </style>\n\n{{#if (slot item @root.slots.Tags)}}\n\t<div class="tags">\n\t\t<pnp-icon data-name="Tag" aria-hidden="true" data-theme-variant="{{JSONstringify @root.theme}}"></pnp-icon>\n\t\t<div>\n\t\t\t{{#each (split (slot item @root.slots.Tags) ",") as |tag| }}\n\t\t\t\t<span>{{trim tag}}</span>\n\t\t\t{{/each}}\n\t\t</div>\n\t</div>\n{{/if}}`, useHandlebarsExpr: true, supportHtml: true },
                                                    { name: strings.Layouts.Cards.Fields.PreviewImage, field: 'previewImage',  value: "{{slot item @root.slots.PreviewImageUrl}}", useHandlebarsExpr: true, supportHtml: false },
                                                    { name: strings.Layouts.Cards.Fields.PreviewUrl, field: 'previewUrl' , value: "{{slot item @root.slots.PreviewUrl}}", useHandlebarsExpr: true, supportHtml: false },
                                                    { name: strings.Layouts.Cards.Fields.Date, field: 'date', value: "{{getDate (slot item @root.slots.Date) 'LL'}}", useHandlebarsExpr: true, supportHtml: false },
                                                    { name: strings.Layouts.Cards.Fields.Url, field: 'href', value: '{{slot item @root.slots.PreviewUrl}}', useHandlebarsExpr: true, supportHtml: false },
                                                    { name: strings.Layouts.Cards.Fields.Author, field: 'author', value: "{{slot item @root.slots.Author}}", useHandlebarsExpr: true, supportHtml: false },
                                                    { name: strings.Layouts.Cards.Fields.ProfileImage, field: 'profileImage', value: "/_layouts/15/userphoto.aspx?size=L&username={{getUserEmail (slot item @root.slots.UserEmail)}}", useHandlebarsExpr: true, supportHtml: false  },
                                                    { name: strings.Layouts.Cards.Fields.FileExtension, field: 'fileExtension', value: "{{slot item @root.slots.FileType}}", useHandlebarsExpr: true, supportHtml: false },
                                                    { name: strings.Layouts.Cards.Fields.IsContainer, field: 'isContainer', value: "{{slot item @root.slots.IsFolder}}", useHandlebarsExpr: true, supportHtml: false }
                                                ] as IComponentFieldsConfiguration[];
            
        this.properties.isCompactCard = this.properties.isCompactCard !== null && this.properties.isCompactCard !== undefined ?  this.properties.isCompactCard: false;
        this.properties.showFileIconCard = this.properties.showFileIconCard !== null && this.properties.showFileIconCard !== undefined ?  this.properties.showFileIconCard: true;
        this.properties.enablePreview = this.properties.enablePreview !== null && this.properties.enablePreview !== undefined ?  this.properties.enablePreview: false; // Watch out performance issues if too many items displayed
        this.properties.preferedCardNumberPerRow = this.properties.preferedCardNumberPerRow ? this.properties.preferedCardNumberPerRow : 3;
        this.properties.columnSizePercentage = this.properties.columnSizePercentage ? this.properties.columnSizePercentage : 33;

        const { PropertyFieldToggleWithCallout } = await import(
            /* webpackChunkName: 'pnp-modern-search-results-cards-layout' */
            '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout'
        );

        const { CalloutTriggers } = await import(
            /* webpackChunkName: 'pnp-modern-search-results-cards-layout' */
            '@pnp/spfx-property-controls/lib/common/callout/Callout'
        );
        
        this._propertyFieldCollectionData = PropertyFieldCollectionData;
        this._customCollectionFieldType = CustomCollectionFieldType;
        this._propertyFieldToogleWithCallout = PropertyFieldToggleWithCallout;
        this._propertyFieldCalloutTriggers = CalloutTriggers;
    }

    public getPropertyPaneFieldsConfiguration(availableFields: string[], dataContext?: IDataContext): IPropertyPaneField<any>[] {

        let availableOptions: IComboBoxOption[] = availableFields.map((fieldName) => { return { key: fieldName, text: fieldName } as IComboBoxOption; });
        let sortableFields = [];
        if (dataContext) {
            sortableFields = dataContext.sorting.selectedSortableFields.map(field => {
                return {
                    key: field,
                    text: field,
                } as IComboBoxOption;
            });
        }

        // Sort ascending
        availableOptions = availableOptions.sort((a, b) => {

            const aValue = a.text ? a.text : a.key ? a.key.toString() : null;
            const bValue = b.text ? b.text : b.key ? b.key.toString() : null;

            if (aValue && bValue) {
                if (aValue.toLowerCase() > bValue.toLowerCase()) return 1;
                if (bValue.toLowerCase() > aValue.toLowerCase()) return -1;
            }

            return 0;
        });

        // Column builder
        let propertyPaneFields: IPropertyPaneField<any>[] = [

            PropertyPaneDropdown('layoutProperties.SelectedView', {
                label: "Selected View",
                options: [{
                    type: PropertyPaneDropdownOptionType.Normal,
                    key: "Details List",
                    text: "Details List",
                } as IPropertyPaneDropdownOption,
                {
                    key: "Simple List",
                    text: "Simple List"
                } as IPropertyPaneDropdownOption],
                selectedKey: this.properties.SelectedView
            }),

            PropertyPaneTextField('layoutProperties.responsiveWidth', {
                label: "Responsive Width (px)",
                value: this.properties.responsiveWidth
            }),
            PropertyPaneHorizontalRule(),

            this._propertyFieldCollectionData('layoutProperties.detailsListColumns', {
                manageBtnLabel: strings.Layouts.DetailsList.ManageDetailsListColumnLabel,
                key: 'layoutProperties.detailsListColumns',
                panelHeader: strings.Layouts.DetailsList.ManageDetailsListColumnLabel,
                panelDescription: strings.Layouts.DetailsList.ManageDetailsListColumnDescription,
                enableSorting: true,
                label: strings.Layouts.DetailsList.ManageDetailsListColumnLabel,
                value: this.properties.detailsListColumns,
                fields: [
                    {
                        id: 'name',
                        title: strings.Layouts.DetailsList.DisplayNameColumnLabel,
                        type: this._customCollectionFieldType.string,
                        required: true,
                    },
                    {
                        id: 'value',
                        title: strings.Layouts.DetailsList.ValueColumnLabel,
                        type: this._customCollectionFieldType.custom,
                        required: true,
                        onCustomRender: (field, value, onUpdate, item, itemId, onCustomFieldValidation) => {
                            return React.createElement("div", { key: `${field.id}-${itemId}` },
                                React.createElement(TemplateValueFieldEditor, {
                                    currentItem: item,
                                    field: field,
                                    useHandlebarsExpr: item.useHandlebarsExpr,
                                    onUpdate: onUpdate,
                                    value: value,
                                    availableProperties: availableOptions,
                                } as ITemplateValueFieldEditorProps)
                            );
                        }
                    },
                    {
                        id: 'useHandlebarsExpr',
                        type: this._customCollectionFieldType.boolean,
                        defaultValue: false,
                        title: strings.Layouts.DetailsList.UseHandlebarsExpressionLabel
                    },
                    {
                        id: 'minWidth',
                        title: strings.Layouts.DetailsList.MinimumWidthColumnLabel,
                        type: this._customCollectionFieldType.number,
                        required: false,
                        defaultValue: 50
                    },
                    {
                        id: 'maxWidth',
                        title: strings.Layouts.DetailsList.MaximumWidthColumnLabel,
                        type: this._customCollectionFieldType.number,
                        required: false,
                        defaultValue: 310
                    },
                    {
                        id: 'enableSorting',
                        title: strings.Layouts.DetailsList.SortableColumnLabel,
                        type: this._customCollectionFieldType.boolean,
                        defaultValue: false,
                        required: false
                    },
                    {
                        id: 'valueSorting',
                        title: strings.Layouts.DetailsList.ValueSortingColumnLabel,
                        type: this._customCollectionFieldType.custom,
                        onCustomRender: (field, _value, onUpdate, item, itemId, onCustomFieldValidation) => {
                            return React.createElement("div", { key: `${field.id}-${itemId}` },
                                React.createElement(AsyncCombo, {
                                    allowFreeform: false,
                                    availableOptions: sortableFields,
                                    placeholder: !item.valueSorting && sortableFields.length > 0 ? strings.Layouts.DetailsList.ValueSortingColumnLabel : strings.Layouts.DetailsList.ValueSortingColumnNoFieldsLabel,
                                    textDisplayValue: item[field.id] ? item[field.id] : '',
                                    defaultSelectedKey: item[field.id] ? item[field.id] : '',
                                    disabled: !item.enableSorting,
                                    onUpdate: (filterValue: IComboBoxOption) => {
                                        onUpdate(field.id, filterValue.key);
                                    }
                                } as IAsyncComboProps));
                        }
                    },
                    {
                        id: 'isResizable',
                        title: strings.Layouts.DetailsList.ResizableColumnLabel,
                        type: this._customCollectionFieldType.boolean,
                        defaultValue: false,
                        required: false
                    },
                    {
                        id: 'isMultiline',
                        title: strings.Layouts.DetailsList.MultilineColumnLabel,
                        type: this._customCollectionFieldType.boolean,
                        defaultValue: false,
                        required: false
                    }
                    
                ]
            }),
            PropertyPaneButton('layoutProperties.resetFields', {
                buttonType: PropertyPaneButtonType.Command,
                icon: 'Refresh',
                text: strings.Layouts.DetailsList.ResetFieldsBtnLabel,
                onClick: () => {
                    // Just reset the fields
                    this.properties.detailsListColumns = null;
                    this.onInit();
                }
            }),
            // Compact mode
            PropertyPaneToggle('layoutProperties.isCompact', {
                label: strings.Layouts.DetailsList.CompactModeLabel,
                checked: this.properties.isCompact ? this.properties.isCompact : true
            }),
            PropertyPaneToggle('layoutProperties.showFileIcon', {
                label: strings.Layouts.DetailsList.ShowFileIcon,
                checked: this.properties.showFileIcon
            }),

            // SIMPLE LIST
            PropertyPaneHorizontalRule(),
            PropertyPaneToggle('layoutProperties.showFileIconList', {
                label: strings.Layouts.SimpleList.ShowFileIconLabel
            }),
            PropertyPaneToggle('layoutProperties.showItemThumbnail', {
                label: strings.Layouts.SimpleList.ShowItemThumbnailLabel
            }),

            // CARDS
            PropertyPaneHorizontalRule(),
            // Careful, the property names should match the React components props. These will be injected in the Handlebars template context and passed as web component attributes
            this._propertyFieldCollectionData('layoutProperties.documentCardFields', {
                manageBtnLabel: strings.Layouts.Cards.ManageTilesFieldsLabel,
                key: 'layoutProperties.documentCardFields',
                panelHeader: strings.Layouts.Cards.ManageTilesFieldsLabel,
                panelDescription: strings.Layouts.Cards.ManageTilesFieldsPanelDescriptionLabel,
                enableSorting: false,
                disableItemCreation: true,
                disableItemDeletion: true,
                label: strings.Layouts.Cards.ManageTilesFieldsLabel,
                value: this.properties.documentCardFields,
                fields: [
                    {
                        id: 'name',
                        type: this._customCollectionFieldType.string,
                        disableEdit: true,
                        title: strings.Layouts.Cards.PlaceholderNameFieldLabel
                    },
                    {
                        id: 'supportHtml',
                        type: this._customCollectionFieldType.custom,
                        disableEdit: true,
                        title: strings.Layouts.Cards.SupportHTMLColumnLabel,
                        onCustomRender: (field, value, onUpdate, item, itemId, onCustomFieldValidation) => {
                            if (item.supportHtml) {
                                return React.createElement(Icon, { iconName: 'CheckMark' } as IIconProps);
                            }
                        }
                    },
                    {
                        id: 'value',
                        title: strings.Layouts.Cards.PlaceholderValueFieldLabel,
                        type: this._customCollectionFieldType.custom,
                        required: true,
                        onCustomRender: (field, value, onUpdate, item, itemId, onCustomFieldValidation) => {
                            return React.createElement("div", { key: `${field.id}-${itemId}` }, 
                                React.createElement(TemplateValueFieldEditor, {
                                    currentItem: item,
                                    field: field,
                                    useHandlebarsExpr: item.useHandlebarsExpr,
                                    onUpdate: onUpdate,
                                    value: value,
                                    availableProperties: availableOptions,
                                } as ITemplateValueFieldEditorProps)
                            );
                        }
                    },
                    {
                        id: 'useHandlebarsExpr',
                        type: this._customCollectionFieldType.boolean,
                        title: strings.Layouts.Cards.UseHandlebarsExpressionLabel
                    }
                ]
            }),
            PropertyPaneButton('layoutProperties.resetFields', {
                buttonType: PropertyPaneButtonType.Command,
                icon: 'Refresh',
                text: strings.Layouts.Cards.ResetFieldsBtnLabel,
                onClick: ()=> {
                    // Just reset the fields
                    this.properties.documentCardFields = null;
                    this.onInit();
                }
            }),
            this._propertyFieldToogleWithCallout('layoutProperties.enablePreview', {
                label: strings.Layouts.Cards.EnableItemPreview,
                calloutTrigger: this._propertyFieldCalloutTriggers.Hover,
                key: 'layoutProperties.enablePreview',
                calloutContent: React.createElement('p', { style:{ maxWidth: 250, wordBreak: 'break-word' }}, strings.Layouts.Cards.EnableItemPreviewHoverMessage),
                onText: strings.General.OnTextLabel,
                offText: strings.General.OffTextLabel,
                checked: this.properties.enablePreview
            }),
            PropertyPaneToggle('layoutProperties.showFileIconCard', {
                label: strings.Layouts.Cards.ShowFileIcon,
                checked: this.properties.showFileIconCard
            }),
            PropertyPaneToggle('layoutProperties.isCompactCard', {
                label: strings.Layouts.Cards.CompactModeLabel,               
                checked: this.properties.isCompactCard
            }),
            PropertyPaneSlider('layoutProperties.preferedCardNumberPerRow', {
                label: strings.Layouts.Cards.PreferedCardNumberPerRow,
                min: 1,
                max: 6,
                step: 1,
                showValue: true,
                value: this.properties.preferedCardNumberPerRow,                
            }),
            PropertyPaneHorizontalRule()

        ];
        
        // Show file icon
        if (this.properties.showFileIcon) {

            propertyPaneFields.push(
                PropertyPaneDropdown('layoutProperties.fieldIconExtension', {
                    label: strings.Layouts.DetailsList.FileExtensionFieldLabel,
                    options: availableOptions,
                    selectedKey: this.properties.fieldIconExtension
                })
            );
        }

        propertyPaneFields.push(
            PropertyPaneToggle('layoutProperties.enableGrouping', {
                label: strings.Layouts.DetailsList.EnableGrouping,
                checked: this.properties.enableGrouping
            }));

        // Grouping options
        if (this.properties.enableGrouping) {
            propertyPaneFields.push(
                PropertyPaneDropdown('layoutProperties.groupByField', {
                    label: strings.Layouts.DetailsList.GroupByFieldLabel,
                    options: availableOptions,
                    selectedKey: this.properties.groupByField,
                }),
                this._propertyPaneWebPartInformation({
                    description: `<small>${strings.Layouts.DetailsList.GroupingDescription}</small>`,
                    key: 'queryText'
                }),
                PropertyPaneToggle('layoutProperties.groupsCollapsed', {
                    label: strings.Layouts.DetailsList.CollapsedGroupsByDefault,
                    checked: this.properties.groupsCollapsed
                })
            );
        }

        return propertyPaneFields;
    }

    // The newValue and oldValue values got swapped somehow... or something else is going on
    public onPropertyUpdate(propertyPath: string, newValue: any, oldValue: any) {

        if (propertyPath.localeCompare('layoutProperties.enableGrouping') === 0) {
            this.properties.groupByField = '';
        }
        if (propertyPath.localeCompare('layoutProperties.preferedCardNumberPerRow') === 0) {
            //// console.log("preferedCardNumberPerRow update!", newValue, oldValue, this.properties.columnSizePercentage = Math.floor(100 /newValue)-1);

            // Calculate the correct % for card flex-basis
            this.properties.columnSizePercentage = Math.floor(100 /oldValue)-1; 
        }
        if (propertyPath.localeCompare('layoutProperties.responsiveWidth') === 0) {
            if(oldValue === "") this.properties.responsiveWidth = "640";
        }

    }
}
