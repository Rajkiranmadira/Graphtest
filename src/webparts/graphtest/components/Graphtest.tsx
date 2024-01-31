import * as React from 'react';

import { IGraphtestProps } from './IGraphtestProps';

import {MSGraphClientV3} from '@microsoft/sp-http';
import {DetailsList, PrimaryButton, Stack, StackItem} from 'office-ui-fabric-react'

import * as XLSX from 'xlsx';  
import {saveAs}  from 'file-saver';

import { sp } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"; 
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import Comp1 from './Comp1/Comp1';
import { HashRouter, Route, Routes } from 'react-router-dom';
import Comp2 from './Comp2/Comp2';
import Comp3 from './Comp3/Comp3';
import Comp4 from './Comp4/Comp4';
// import { version } from 'react-dom';

export interface IListItem {
  id: number;
  title: string;
}



export interface IUser{
  displayName:string;
  mail:string
}

export interface ISubSite{
  webUrl:string;
  displayName:string;
}

export interface ISitesLoop{
  siteNames:string;
}

export interface IUserState {
  userState:IUser[];
  subSiteData: ISubSite[];
  updateText:IListItem[];
  siteNamesArray:ISitesLoop[];
}


//excel

const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';  
const fileExtension = '.xlsx';  
var Heading = [  ["ID","Title"],];  
const saveExcel = (ListData:any) => {  
  if(ListData.length>0)  
  {  
  const ws = XLSX.utils.book_new();  
   // const ws = XLSX.utils.json_to_sheet(csvData,{header:["A","B","C","D","E","F","G"], skipHeader:false});  
   XLSX.utils.sheet_add_aoa(ws, Heading);  
   XLSX.utils.sheet_add_json(ws, ListData, { origin: 'A2', skipHeader: true });         
    const wb = { Sheets: { 'data': ws }, SheetNames: ['data'] };  
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });  
    const data = new Blob([excelBuffer], {type: fileType});  
    saveAs(data, 'Data' + fileExtension);  
  }  
}  


export default class Graphtest extends React.Component<IGraphtestProps,IUserState,ISitesLoop> {

  public alluser:IUser[]=[];
  public subSiteData:ISubSite[]=[];
  public web = Web("https://tqkvy.sharepoint.com/sites/Rajsite/Sitea");  
  public web1 = Web("https://tqkvy.sharepoint.com/sites/Rajsite/Siteb");  
  public sitesNamesArray:ISitesLoop[]=[];
  

  constructor(props:IGraphtestProps){
    super(props)
    this.state={userState:[],subSiteData:[],updateText:[],siteNamesArray:[]}

    sp.setup({
      spfxContext: this.props.context
    });
    this.Listdata=this.Listdata.bind(this); 
  }

  public realGetListNames = (siteNamesArray:ISitesLoop[]) =>{
    console.log("realGetListNames");
    let subsitenames:string="";
    let items1: IListItem[] = [];

    siteNamesArray.forEach(element => {
      subsitenames=element.siteNames;
      
      console.log(subsitenames+"elementtttttttt")

      
  
  let webu = Web(subsitenames+"");  
    webu.lists.getByTitle("DLS").items().then((items) => {  
    let hi:any = items;
    
    items.forEach(element => {
      items1.push({ id: element.Id, title: element.Title });
    });
    this.setState({  
      updateText: items1  
  });  
    
    console.log(this.state.updateText+"royyyyyyy")
    console.log(hi+"royyyyyyyy")
}).catch((err) => {  
    console.log(err);  
}); 


    });
    

    

    // siteNamesArray.map((result:any) => {
    //   this.subSiteData.push({displayName:result.displayName,
    //    webUrl:result.webUrl
      
    //  });
    // })

  }

  componentDidMount(): void {
    this.setState({subSiteData:[]})
    this.subSiteData=[];
  this.props.context.msGraphClientFactory.
  getClient('3').then((msGraphClient:MSGraphClientV3)=>{
    msGraphClient.api("sites/tqkvy.sharepoint.com,19ea9d46-a9a3-446f-82a7-cc6dac9aa44b,c8250e37-fdf8-4d1c-93a1-f0a0899bef62/sites").
    version("v1.0").
    get((err,res) =>{
      if(err){
        console.log("error occured",err)
      }
      console.log(res+"ressssssssssssssss")
      res.value.map((result:any) => {
         this.subSiteData.push({displayName:result.displayName,
          webUrl:result.webUrl
         
        });

        //putting inti an array ignore this
        this.sitesNamesArray.push({siteNames:result.webUrl})
        console.log(this.sitesNamesArray+"alludu")
        
      })
      this.setState({siteNamesArray:this.sitesNamesArray});
      this.realGetListNames(this.sitesNamesArray);

      
      

      this.setState({subSiteData:this.subSiteData});
    });


  })

  //get list items
  
//   let items1: IListItem[] = [];
//   this.web.lists.getByTitle("DLS").items().then((items) => {  
//     let hi:any = items;
    
//     items.forEach(element => {
//       items1.push({ id: element.Id, title: element.Title });
//     });
//     this.setState({  
//         updateText: items1  
//     });  
//     console.log(this.state.updateText+"hiiiiiiiiiiiii")
//     console.log(hi+"hiiiiiiiiiiiii")
// }).catch((err) => {  
//     console.log(err);  
// }); 


//2list 
// this.web1.lists.getByTitle("DLS").items().then((itemsb) => {  
//   let hi:any = itemsb;
  
//   itemsb.forEach(element => {
//     items1.push({ id: element.Id, title: element.Title });
//   });
//   this.setState({  
//       updateText: items1  
//   });  
//   console.log(this.state.updateText+"hiiiiiiiiiiiii")
//   console.log(hi+"hiiiiiiiiiiiii")
// }).catch((err) => {  
//   console.log(err);  
// }); 




    
  }



  //Save excel
  private  Listdata=async ()=>{    
       

        saveExcel(this.state.subSiteData);  
        return this.state.updateText;          
         
    
  } 
  
  //save excel
  



  public GetUsers = () =>{
    this.setState({userState:[]})
    this.alluser=[];
  this.props.context.msGraphClientFactory.
  getClient('3').then((msGraphClient:MSGraphClientV3)=>{
    msGraphClient.api("users").
    version("v1.0").
    select("displayName,mail").get((err,res) =>{
      if(err){
        console.log("error occured",err)
      }

      console.log(res+"ressssss")
      res.value.map((result:any) => {
         this.alluser.push({displayName:result.displayName,
          mail:result.mail
        });
      })
      this.setState({userState:this.alluser})
    });


  })

 

  }


  //subsite

  public GetSubSite = () =>{
    this.setState({subSiteData:[]})
    this.subSiteData=[];
  this.props.context.msGraphClientFactory.
  getClient('3').then((msGraphClient:MSGraphClientV3)=>{
    msGraphClient.api("sites/tqkvy.sharepoint.com,19ea9d46-a9a3-446f-82a7-cc6dac9aa44b,c8250e37-fdf8-4d1c-93a1-f0a0899bef62/sites").
    version("v1.0").
    get((err,res) =>{
      if(err){
        console.log("error occured",err)
      }
      console.log(res+"ressssssssssssssss")
      res.value.map((result:any) => {
         this.subSiteData.push({displayName:result.displayName,
          webUrl:result.webUrl
        });
      })
      this.setState({subSiteData:this.subSiteData})
    });


  })

 

  }


  public render(): React.ReactElement<IGraphtestProps> {
    

    return (
      <div>
        <PrimaryButton text='Search user' onClick={this.GetUsers}></PrimaryButton>
        <PrimaryButton text='Search subsite' onClick={this.GetSubSite}></PrimaryButton>
        <button type='button' onClick={this.Listdata}>Export to Excel</button> 

        <DetailsList items={this.state.userState}></DetailsList>
        <DetailsList items={this.state.subSiteData}></DetailsList>
        <DetailsList items={this.state.updateText}></DetailsList>


        <HashRouter>
          <Stack horizontal>
            <Comp1></Comp1>
            <StackItem>
              <switch>
                <Routes>
                <Route 
                   path="/comp2"
                   Component={()=><Comp2/>}
                ></Route>
                <Route 
                   path="/comp3"
                   Component={()=><Comp3 description={''} context={this.props.context} siteUrl={this.props.context.pageContext.web.absoluteUrl}/>}
                ></Route>
                <Route 
                   path="/comp4"
                   Component={()=><Comp4/>}
                ></Route>
                </Routes>
              </switch>
            </StackItem>

          </Stack>
        </HashRouter>
        

      </div>
    );
  }
}
