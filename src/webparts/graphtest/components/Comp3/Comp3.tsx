import * as React from 'react';

import { IComp3Props } from './IComp3Props';
import {sp,Web} from '@pnp/sp/presets/all';
import { IGraphtestProps } from '../IGraphtestProps';
import { Icomp2State } from '../Icomp2State';
import { TextField,PrimaryButton, Toggle } from '@fluentui/react';

import '@pnp/sp/lists';
import '@pnp/sp/items'




// import { version } from 'react-dom';




export default class Graphtest extends React.Component<IGraphtestProps,Icomp2State> {



constructor(props:IGraphtestProps,state:Icomp2State){
  super(props);
  sp.setup({
    spfxContext:this.props.context
  });
  this.state={
    Title:"",
    PhoneNumber:'',
    Married:false,
    HTML:[]
  }
}


async componentDidMount() {
  await this.getPeoplePicker();
}

public async getPeoplePicker(){
  let web1= Web(this.props.siteUrl);

  const items1: any[] = await web1.lists.getByTitle("Creds").items.select('*','Approver/Title').expand('Approver').getAll();

  console.log("comp3")
  console.log(items1);

  

  var tabledata = <table >
  <thead>
    <tr>
      <th>Approver1</th>
      <th>Approver2</th>
      
    </tr>
  </thead>
  <tbody>
    {items1 && items1.map((item:any, i:any) => {
      var myapproverlenght:Number=item.Approver.length;
      console.log(myapproverlenght)
      return [
        <tr key={i}>

         {item.Approver.map((it:any, i1:any)=>{
           return[
            <td>{it.Title}</td>
           ]

         }  )}
        

          
         


          
          {/* <td>{item.Approver[i].Title}</td>
          <td>{item.Approver[i+1].Title}</td> */}
          
          
        </tr>
      ];
    })}
  </tbody>

</table>;

this.setState({HTML:tabledata});

}




 public async saveDate(){

  let web= Web(this.props.siteUrl);

  await web.lists.getByTitle("Creds").items.add({
    Title:this.state.Title,
    PhoneNumber:this.state.PhoneNumber,
    Married: this.state.Married
  })
  .then((data)=>{
    console.log("No errors"+data);
    this.setState({
      Title:"",
    PhoneNumber:"",
    Married: false
    });
  })
  .catch((err)=>{
    console.error("Error occured")
  })
  alert("Item Created Successfully");

  
  
}

public formEvents = (feildName:keyof Icomp2State,value:string|number|boolean):void =>{

  this.setState({[feildName]:value} as unknown as Pick<Icomp2State, keyof Icomp2State> );

}

 



  public render(): React.ReactElement<IComp3Props> {
    

    return (
      <div>
        <h1>Welcome to Component3</h1>

        <TextField label='UserName' value={this.state.Title} 
          onChange={(_,title)=>this.formEvents("Title",title||'')}/>

        <TextField label='PhoneNumber' value={this.state.PhoneNumber}
           onChange={(_,PhoneNumber) => this.formEvents("PhoneNumber",parseInt(PhoneNumber || '0'))}/>

        <Toggle label='Married' checked={this.state.Married}
        onChange={(checked)=> this.formEvents("Married",!!checked)}/>

        <PrimaryButton onClick={()=>this.saveDate()} iconProps={{iconName:'save'}}/>

        {this.state.HTML}

      </div>
    );
  }
}
