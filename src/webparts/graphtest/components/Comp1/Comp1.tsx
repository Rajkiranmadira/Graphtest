import * as React from 'react';

import { IComp1Props } from './IComp1Props';


import { Nav,INavLinkGroup } from 'office-ui-fabric-react';



const group:INavLinkGroup[]=[
  {
    links:[
      {name:"Component2",url:"#/comp2"},
      {name:"Component3",url:"#/comp3"},
      {name:"Component4",url:"#/comp4"},
    ]
  }
 ]

export default class Graphtest extends React.Component<IComp1Props,{}> {







  public render(): React.ReactElement<IComp1Props> {
    

    return (

      <Nav groups={group}></Nav>
      

    );
  }
}
