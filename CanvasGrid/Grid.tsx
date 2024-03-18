import {
    DetailsList,
    ConstrainMode,
    DetailsListLayoutMode,
    IColumn,
    IDetailsHeaderProps,
} from '@fluentui/react/lib/DetailsList';
import { Overlay } from '@fluentui/react/lib/Overlay';
import {
    ScrollablePane,
    ScrollbarVisibility
} from '@fluentui/react/lib/ScrollablePane';
import { Stack } from '@fluentui/react/lib/Stack';
import { Sticky } from '@fluentui/react/lib/Sticky';
import { StickyPositionType } from '@fluentui/react/lib/Sticky';
import { IObjectWithKey } from '@fluentui/react/lib/Selection';
import { IRenderFunction } from '@fluentui/react/lib/Utilities';
import * as React from 'react';
import { useConst, useForceUpdate } from "@fluentui/react-hooks";
import { Selection } from "@fluentui/react/lib/Selection";
import { SelectionMode } from "@fluentui/react/lib/Utilities";

type DataSet = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord & IObjectWithKey;

export interface GridProps {
    width?: number;
    height?: number;
    columns: ComponentFramework.PropertyHelper.DataSetApi.Column[];
    records: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord>;
    sortedRecordIds: string[];
    hasNextPage: boolean;
    hasPreviousPage: boolean;
    totalResultCount: number;
    currentPage: number;
    sorting: ComponentFramework.PropertyHelper.DataSetApi.SortStatus[];
    filtering: ComponentFramework.PropertyHelper.DataSetApi.FilterExpression;
    resources: ComponentFramework.Resources;
    itemsLoading: boolean;
    highlightValue: string | null;
    highlightColor: string | null;
    setSelectedRecords: (ids: string[]) => void;
    onNavigate: (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) => void;
}

const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
    if (props && defaultRender) {
        return (
            <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
                {defaultRender({
                    ...props,
                })}
            </Sticky>
        );
    }
    return null;
};

const onRenderItemColumn = (
    item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord,
    index?: number,
    column?: IColumn,
) => {
    if (column && column.fieldName && item) {
        return <>{item?.getFormattedValue(column.fieldName)}</>;
    }
    return <></>;
};

export const Grid = React.memo((props: GridProps) => {
    const {
        records,
        sortedRecordIds,
        columns,
        width,
        height,
        hasNextPage,
        hasPreviousPage,
        sorting,
        filtering,
        currentPage,
        itemsLoading,
        setSelectedRecords,
        onNavigate,
    } = props;

    const forceUpdate = useForceUpdate();
    const onSelectionChanged = React.useCallback(() => {
        const items = selection.getItems() as DataSet[];
        const selected = selection.getSelectedIndices().map((index: number) => {
            const item: DataSet | undefined = items[index];
            return item && items[index].getRecordId();
        });

        setSelectedRecords(selected);
        forceUpdate();
    }, [forceUpdate]);

    const selection: Selection = useConst(() => {
        return new Selection({
            selectionMode: SelectionMode.single,
            onSelectionChanged: onSelectionChanged,
        });
    });

    const [isComponentLoading, setIsLoading] = React.useState<boolean>(false);

    const items: (DataSet | undefined)[] = React.useMemo(() => {
        setIsLoading(false);

        const sortedRecords: (DataSet | undefined)[] = sortedRecordIds.map((id) => {
            const record = records[id];
            return record;
        });

        return sortedRecords;
    }, [records, sortedRecordIds, hasNextPage, setIsLoading]);

    const gridColumns = React.useMemo(() => {
        return columns
            .filter((col) => !col.isHidden && col.order >= 0)
            .sort((a, b) => a.order - b.order)
            .map((col) => {
                const sortOn = sorting && sorting.find((s) => s.name === col.name);
                const filtered =
                    filtering &&
                    filtering.conditions &&
                    filtering.conditions.find((f) => f.attributeName == col.name);
                return {
                    key: col.name,
                    name: col.displayName,
                    fieldName: col.name,
                    isSorted: sortOn != null,
                    isSortedDescending: sortOn?.sortDirection === 1,
                    isResizable: true,
                    isFiltered: filtered != null,
                    data: col,
                } as IColumn;
            });
    }, [columns, sorting]);

    const rootContainerStyle: React.CSSProperties = React.useMemo(() => {
        return {
            height: height,
            width: width,
        };
    }, [width, height]);

    return (
        <Stack verticalFill grow style={rootContainerStyle}>
            <Stack.Item grow style={{ position: 'relative', backgroundColor: 'white' }}>
                <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                    <DetailsList
                        columns={gridColumns}
                        onRenderItemColumn={onRenderItemColumn}
                        onRenderDetailsHeader={onRenderDetailsHeader}
                        items={items}
                        setKey={`set${currentPage}`} // Ensures that the selection is reset when paging
                        initialFocusedIndex={0}
                        checkButtonAriaLabel="select row"
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        constrainMode={ConstrainMode.unconstrained}
                        selection={selection}
                        onItemInvoked={onNavigate}
                    ></DetailsList>
                </ScrollablePane>
                {(itemsLoading || isComponentLoading) && <Overlay />}
            </Stack.Item>
        </Stack>
    );
});

Grid.displayName = 'Grid';