import { initializeIcons } from "@fluentui/react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { Grid } from "./Grid";

initializeIcons(undefined, { disableWarnings: true });

export class CanvasGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    notifyOutputChanged: () => void;
    container: HTMLDivElement;
    context: ComponentFramework.Context<IInputs>;
    sortedRecordsIds: string[] = [];
    resources: ComponentFramework.Resources;
    isTestHarness: boolean;
    records: {
        [id: string]: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
    };
    currentPage = 1;
    filteredRecordCount?: number;

    /**
     * Empty constructor.
     */
    constructor() {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.container = container;
        this.context = context;
        this.context.mode.trackContainerResize(true);
        this.resources = this.context.resources;
        this.isTestHarness = document.getElementById("control-dimensions") !== null;
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        const dataset = context.parameters.records;
        const paging = context.parameters.records.paging;
        const datasetChanged = context.updatedProperties.indexOf("dataset") > -1;
        const resetPaging =
            datasetChanged &&
            !dataset.loading &&
            !dataset.paging.hasPreviousPage &&
            this.currentPage !== 1;
    
        if (resetPaging) {
            this.currentPage = 1;
        }
        if (resetPaging || datasetChanged || this.isTestHarness) {
            this.records = dataset.records;
            this.sortedRecordsIds = dataset.sortedRecordIds;
        }
    
        // The test harness provides width/height as strings
        const allocatedWidth = parseInt(
            context.mode.allocatedWidth as unknown as string
        );
        const allocatedHeight = parseInt(
            context.mode.allocatedHeight as unknown as string
        );
    
        ReactDOM.render(
            React.createElement(Grid, {
                width: allocatedWidth,
                height: allocatedHeight,
                columns: dataset.columns,
                records: this.records,
                sortedRecordIds: this.sortedRecordsIds,
                hasNextPage: paging.hasNextPage,
                hasPreviousPage: paging.hasPreviousPage,
                currentPage: this.currentPage,
                totalResultCount: paging.totalResultCount,
                sorting: dataset.sorting,
                filtering: dataset.filtering && dataset.filtering.getFilter(),
                resources: this.resources,
                itemsLoading: dataset.loading,
                highlightValue: this.context.parameters.HighlightValue.raw,
                highlightColor: this.context.parameters.HighlightColor.raw,
            }),
            this.container
        );
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this.container);
    }

}
