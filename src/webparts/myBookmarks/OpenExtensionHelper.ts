import { MSGraphClient } from "@microsoft/sp-client-preview";
import { WebPartContext } from "@microsoft/sp-webpart-base";

class OpenExtensionHelper {
    private _graphClient: MSGraphClient;

    constructor(context:WebPartContext) {
        this._graphClient = context.serviceScope.consume(
            MSGraphClient.serviceKey
        )
    }

    GetOpenExtension(extensionKey) {
          this._graphClient
            .api("me")
            .version("v1.0")
            .select("id,displayName")
            .expand("extensions")
            .get((err, res) => {
              if(err) {
                console.error(err);
                return;
              }
      
              console.log(res);

              
            })  
    } 
}

// class OpenExtensionHelper {

//     //private var _graphClient : MSGraphClient = null;

//     // private constructor(context:WebPartContext){
        
//     // }
      
//     // function GetExtension(extensionKey){
  
      
//     //       graphClient
//     //         .api("me")
//     //         .version("v1.0")
//     //         .select("id,displayName")
//     //         .expand("extensions")
//     //         .get((err, res) => {
//     //           if(err) {
//     //             console.error(err);
//     //             return;
//     //           }
      
//     //           console.log(res);
//     //         })
//     // }

// }