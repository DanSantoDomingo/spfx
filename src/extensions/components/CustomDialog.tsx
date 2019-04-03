import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { IPeriodYearItem } from "../closePeriod/ClosePeriodCommandSet";
import { sp, ItemAddResult } from "@pnp/pnpjs";
import {
    Label,
    autobind,
    PrimaryButton,
    Button,
    DialogFooter,
    DialogContent,
    Spinner,
    SpinnerType,
    ProgressIndicator,
} from 'office-ui-fabric-react';

export interface ICustomDialogProps {
    message: string;
    close: () => void;
    selectedItem: IPeriodYearItem;
    //performTask: () => void;
}

export interface ICustomDialogState {
    updatingItem: boolean;
    closingActiveResourceProjection: boolean;
    copyingProjection: boolean;
    askingQuestion: boolean;
    showButtons: boolean;
    showCloseButton: boolean;
    showCompletionMessage: boolean;
    taskProgress: number;
    taskDescription: string;
}


class CustomDialogContent extends React.Component<ICustomDialogProps, ICustomDialogState, {}> {
    private _selectedItem: IPeriodYearItem;
    private _progress: number = 0;
    private _totalWork: number = 0;
    private _barHeight: number = 4;


    private taskProgress(): void {
        this._progress += 1;
       
        //console.log(increase);
        this.setState({
            taskDescription: this._progress + "/" + this._totalWork,
            taskProgress: this._progress / this._totalWork,
        });
    }

    constructor(props: ICustomDialogProps) {
        super(props);
        this._selectedItem = props.selectedItem;
        this.state = {
            updatingItem: false,
            copyingProjection: false,
            closingActiveResourceProjection: false,
            askingQuestion: true,
            showButtons: true,
            showCloseButton: false,
            showCompletionMessage: false,
            taskProgress: 0,
            taskDescription: "0/0",
        };
    }

    public render(): JSX.Element {      
        return <DialogContent title='Update Resource Projection' onDismiss={this.props.close} >
            
            {this.state.askingQuestion && <Label>{this.props.message}</Label>}
            {this.state.updatingItem && <ProgressIndicator label="Updating Period Year..." />}
            {this.state.closingActiveResourceProjection && <ProgressIndicator barHeight={this._barHeight} label="Closing all active Resource Projection..." description={this.state.taskDescription} percentComplete={this.state.taskProgress} />}
            {this.state.copyingProjection && <ProgressIndicator barHeight={this._barHeight} label="Copying Resource Projection from previous quarter.." description={this.state.taskDescription} percentComplete={this.state.taskProgress} />}
            {this.state.showCompletionMessage && <Label>Done.</Label>}
            <DialogFooter>
                {this.state.showButtons && <PrimaryButton text='OK' title='OK' onClick={this.performTask} />}
                {this.state.showButtons && <Button text='Cancel' title='Cancel' onClick={this.props.close} />}
                {this.state.showCloseButton && <Button text='Close' title='Close' onClick={this.props.close} />}
            </DialogFooter>
        </DialogContent>;
    }

    @autobind
    public performTask(): void {
        //console.log("Starting Task..");
        this.closePrevPeriod();
        //this.updateSelectedItem(this._selectedItem);
        //this.getResourceProjItems(this.prevQ(this._selectedItem.Quarter), this.prevYear(this._selectedItem.Quarter, this._selectedItem.Period));
        //this.copyResourceProjItems(this.prevQ(this._selectedItem.Quarter), this.prevYear(this._selectedItem.Quarter, this._selectedItem.Period));
    }

    private prevYear(quarter: number, year: number): string {
        if (quarter == 1) {
            let prev = year - 1;
            return prev.toString();
        }
        else {
            return year.toString();
        }
    }

    private prevQ(quarter: number): string {
        if (quarter == 1) {
            return '4';
        }
        else {
            let prev = quarter - 1;
            return prev.toString();
        }
    }

    private closePrevPeriod(): void {
        this.setState({
            showButtons: false,
            updatingItem: true,
            askingQuestion: false,
        });

        console.log("Updating period list...");
        let list = sp.web.lists.getByTitle("Period Year");
        list.items.filter("Status eq 'O'").get().then((items: IPeriodYearItem[]) => {
            if (items.length > 0) {
                list.items.getById(items[0].Id).update({
                    Status: "C",
                }).then(() => {
                    this.updateSelectedItem(this._selectedItem);
                });
            }
        });
    }

    private updateSelectedItem(item: IPeriodYearItem): void {
        sp.web.lists.getByTitle("Period Year").items.getById(item.Id).update({
            Status: "O"
        }).then(() => {
            this.getResourceProjItems(this.prevQ(item.Quarter), this.prevYear(item.Quarter, item.Period));
        });
    }

    private getResourceProjItems(currentQ: any, year: any): void {
        this.setState({
            updatingItem: false,
            closingActiveResourceProjection: true,
        });

        //console.log("Closing all resource projection of QUARTER: " + this.prevQ(this._selectedItem.Quarter) + " YEAR: " + this.prevYear(this._selectedItem.Quarter, this._selectedItem.Period));
        let list = sp.web.lists.getByTitle("Resource Projection");
        list.items.filter(`Quarter_x0020_Num eq '${currentQ}' & and Period_x0020_Year eq '${year}'`).getAll().then((items) => {
            console.log(items.length);
            this._totalWork = items.length;
            this._progress = 0;
            if (items.length > 0) {
                items.forEach((item) => {
                    list.items.getById(item.Id).update({
                        Status: "Closed",
                    }).then(() => {
                        this.taskProgress();
                        if (this._progress >= this._totalWork) {
                            this.copyResourceProjItems(this.prevQ(this._selectedItem.Quarter), this.prevYear(this._selectedItem.Quarter, this._selectedItem.Period));
                        }
                    });
                });
            }
        });
    }

    private copyResourceProjItems(currentQ: any, year: any): void {
        this._totalWork = 0;
        this._progress = 0;

        this.setState({
            closingActiveResourceProjection: false,
            copyingProjection: true,
            taskDescription: this._progress + "/" + this._totalWork,
            taskProgress: this._progress / this._totalWork,
        });

        //console.log("Copying all resource projection of QUARTER: " + this.prevQ(this._selectedItem.Quarter) + " YEAR: " + this.prevYear(this._selectedItem.Quarter, this._selectedItem.Period));
        let list = sp.web.lists.getByTitle("Resource Projection");
        list.items.filter(`Quarter_x0020_Num eq '${currentQ}' & and Period_x0020_Year eq '${year}'`).getAll().then((items) => {
            if (items.length > 0) {
                this._totalWork = items.length;
                this._progress = 0;
                items.forEach((item) => {
                    list.items.add({
                        Title: item.Title,
                        First: item.First,
                        Last: item.Last,
                        Team: item.Team,
                        Main_x0020_System: item.Main_x0020_System,
                        Sub_x0020_System: item.Sub_x0020_System,
                        Percentage: item.Percentage,
                        Category_x0020_not_x0020_in_x002: item.Category_x0020_not_x0020_in_x002,
                        Group: item.Group,
                        Status: 'Draft',
                        CategoryId: item.CategoryId,
                        EmailId: item.EmailId,
                        Period_x0020_Year: this._selectedItem.Period,
                        Quarter_x0020_Num: this._selectedItem.Quarter,
                    }).then(() => {
                        this.taskProgress();
                        if (this._progress >= this._totalWork) {
                            this.setState({
                                copyingProjection: false,
                                showCloseButton: true,
                                showCompletionMessage: true,
                            });
                        }
                    });
                });

            }

        });
    }
}

export default class CustomDialog extends BaseDialog {
    public message: string;
    public selectedItem: IPeriodYearItem;

    constructor() {
        super({ isBlocking: true });
    }

    public render(): void {
        ReactDOM.render(<CustomDialogContent close={this.close} message={this.message} selectedItem={this.selectedItem}  />, this.domElement);
    }

    protected onAfterClose(): void {
        super.onAfterClose();
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
}
