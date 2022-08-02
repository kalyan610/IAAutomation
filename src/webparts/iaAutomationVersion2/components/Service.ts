import {sp} from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";

export default class Service {

    public mysitecontext: any;

    public constructor(siteUrl: string, Sitecontext: any) {
        this.mysitecontext = Sitecontext;


        sp.setup({
            sp: {
                baseUrl: "https://capcoinc.sharepoint.com/sites/capcointernalapplications/"

                //baseUrl: "https://capcoinc.sharepoint.com/sites/IAAutomationEnvironment/"

            },
        });

    }

    public async getUserByLogin(LoginName:string):Promise<any>{
        try{
            const user = await sp.web.siteUsers.getByLoginName(LoginName).get();
            return user;
        }catch(error){
            console.log(error);
        }
    }


    public async GetListNameandURL():Promise<any>
    {
 
     return await sp.web.lists.getByTitle("URLandListname").items.select('Title','URL','ColName').expand().get().then(function (data) {
 
     return data;
 
     });
 
    }



    public async getCurrentUser(): Promise<any> {
        try {
            return await sp.web.currentUser.get().then(result => {
                return result;
            });
        } catch (error) {
            console.log(error);
        }
      }

      


    private async onDrop(MyColnameDescReq:string,MyListTitle:string,MyRequestorVal:string,MySystemval:string,MyTitleVal:string,MyDescVal:string,MyServiceNowStatus:string,MyServiceNowTicketNumber:string,MyRequestDate:string,MyRequestDelDate:string,MyServDate:string,MyLoginName:string,MyServiceNowText:string,acceptedFiles):Promise<any>     {       
        let Myval='Completed';

        
       if(MyColnameDescReq=='Descriptions')
       {
       
        try
        {
    
        let Filemal=[];
    
        let file=acceptedFiles;
    
        let Varmyval= await sp.web.lists.getByTitle(MyListTitle).items.add({
        
        Title:MyTitleVal,
        Requestor:MyRequestorVal,
        System:MySystemval,
        Descriptions:MyDescVal,
        Did_x0020_this_x0020_start_x0020:MyServiceNowStatus,
        ServiceNowTicket:MyServiceNowTicketNumber,
        Request_x0020_Date:MyRequestDate,
        End:MyRequestDelDate,
        Date_x0020_ServiceNow_x0020_Tick:MyServDate,
        Submitter:MyLoginName,
        ServiceNowRequester:MyServiceNowText,
        ReadytoAction:'No'
        
        }).then (async r => {
          // this will add an attachment to the item we just created to push t sharepoint list
    
        for(var count=0;count<file.length;count++)
        {
         await r.item.attachmentFiles.add(file[count].name, file[count]).then(result => {
        console.log(result);
        
    
          })
    
        }
    
        return Myval;
    
    
    
        })
    
        
    
        return Varmyval;
    
        
      }
    
    
    
      catch (error) {
        console.log(error);
      }

    }

    if(MyColnameDescReq=='Description')
    {

      try
      {
  
      let Filemal=[];
  
      let file=acceptedFiles;
  
      let Varmyval= await sp.web.lists.getByTitle(MyListTitle).items.add({
      
      Title:MyTitleVal,
      Requestor:MyRequestorVal,
      System:MySystemval,
      Description:MyDescVal,
      Did_x0020_this_x0020_start_x0020:MyServiceNowStatus,
      ServiceNowTicket:MyServiceNowTicketNumber,
      Request_x0020_Date:MyRequestDate,
      End:MyRequestDelDate,
      Date_x0020_ServiceNow_x0020_Tick:MyServDate,
      Submitter:MyLoginName,
      ServiceNowRequester:MyServiceNowText,
      ReadytoAction:'No'

      
      }).then (async r => {
        // this will add an attachment to the item we just created to push t sharepoint list
  
      for(var count=0;count<file.length;count++)
      {
       await r.item.attachmentFiles.add(file[count].name, file[count]).then(result => {
      console.log(result);
      
  
        })
  
      }
  
      return Myval;
  
  
  
      })
  
      
  
      return Varmyval;
  
      
    }
  
  
  
    catch (error) {
      console.log(error);
    }


    }
    
    
      
     }



     private async onDrop1(MyColnameDescReq:string,MyListTitle:string,MyRequestorVal:string,MySystemval:string,MyTitleVal:string,MyDescVal:string,MyServiceNowStatus:string,MyRequestDate:string,MyRequestDelDate:string,MyLoginName:string,acceptedFiles):Promise<any>     {       
        let Myval='Completed';
       
        if(MyColnameDescReq=='Description')
        {
        
        try
        {
    
        let Filemal=[];
    
        let file=acceptedFiles;
    
        let Varmyval= await sp.web.lists.getByTitle(MyListTitle).items.add({
        
        Title:MyTitleVal,
        Requestor:MyRequestorVal,
        System:MySystemval,
        Description:MyDescVal,
        Did_x0020_this_x0020_start_x0020:MyServiceNowStatus,
        Request_x0020_Date:MyRequestDate,
        End:MyRequestDelDate,
        Submitter:MyLoginName,
        ReadytoAction:'No'

        
        }).then (async r => {
          // this will add an attachment to the item we just created to push t sharepoint list
    
        for(var count=0;count<file.length;count++)
        {
         await r.item.attachmentFiles.add(file[count].name, file[count]).then(result => {
        console.log(result);
        
    
          })
    
        }
    
        return Myval;
    
    
    
        })
    
        
    
        return Varmyval;
    
        
      }
    
    
    
      catch (error) {
        console.log(error);
      }

    }

    if(MyColnameDescReq=='Descriptions')
    {
      try
      {
  
      let Filemal=[];
  
      let file=acceptedFiles;
  
      let Varmyval= await sp.web.lists.getByTitle(MyListTitle).items.add({
      
      Title:MyTitleVal,
      Requestor:MyRequestorVal,
      System:MySystemval,
      Descriptions:MyDescVal,
      Did_x0020_this_x0020_start_x0020:MyServiceNowStatus,
      Request_x0020_Date:MyRequestDate,
      End:MyRequestDelDate,
      Submitter:MyLoginName,
      ReadytoAction:'No'

      
      }).then (async r => {
        // this will add an attachment to the item we just created to push t sharepoint list
  
      for(var count=0;count<file.length;count++)
      {
       await r.item.attachmentFiles.add(file[count].name, file[count]).then(result => {
      console.log(result);
      
  
        })
  
      }
  
      return Myval;
  
  
  
      })
  
      
  
      return Varmyval;
  
      
    }
  
  
  
    catch (error) {
      console.log(error);
    }


    }
    
    
      
     }


     
     
        

}





















