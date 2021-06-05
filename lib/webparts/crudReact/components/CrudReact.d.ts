import * as React from 'react';
import { ICrudReactProps } from './ICrudReactProps';
import { IReactCRUDState } from './IReactCRUDState';
export default class CrudReact extends React.Component<ICrudReactProps, IReactCRUDState> {
    constructor(props: ICrudReactProps, state: IReactCRUDState);
    render(): React.ReactElement<ICrudReactProps>;
    private getLatestItemId;
    private createItem;
    private readItem;
    private updateItem;
    private deleteItem;
    private handleTitle;
    private handleDesc;
    private AssignedTo;
    private _changeSharing;
    private _getPeoplePickerItems;
    private _log;
    private _onCheckboxChange;
    private validateForm;
    componentDidMount(): Promise<void>;
}
//# sourceMappingURL=CrudReact.d.ts.map