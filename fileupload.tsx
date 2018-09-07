import * as React from 'react'
import { uploadSpreadsheet } from './spreadsheetUpload';

export interface IMainState {
}

export interface IMainProps {}

export class ScoresheetUpload extends React.Component<IMainProps,IMainState> {
    scores:Object


    constructor(){
        super({},{});
        this.scores={}
    }
    handleChange(event:React.ChangeEvent):void{
        let target=event.target as HTMLInputElement
        let file:File=target.files[0]
        this.scores= uploadSpreadsheet(file)
    }
    render() {
        return (
            <div className="fileuploadContainer">
                <input type="file" onChange={this.handleChange}></input>
                <div>{this.scores}</div>
            </div>

        )
    }
}