import { api } from "@/lib/api";

export type BetaFeedbackItem={id:string;user_id:string;category:string;severity:"p0"|"p1"|"p2";status:"open"|"reviewing"|"fixed"|"closed";title:string;description:string;path:string;user_role?:string|null;app_version?:string|null;browser?:string|null;viewport?:string|null;request_id?:string|null;has_attachment:boolean;resolution_notes?:string|null;reporter_name?:string;created_at:string;updated_at:string};

export const betaFeedbackService={
 async create(payload:{category:string;severity:string;title:string;description:string;path:string;app_version?:string;browser?:string;viewport?:string;request_id?:string;screenshot?:File|null}){
  const body=new FormData();
  for(const [k,v] of Object.entries(payload)){if(k!=="screenshot"&&v!=null)body.append(k,String(v));}
  if(payload.screenshot)body.append("screenshot",payload.screenshot);
  return (await api.post("/beta-feedback",body,{headers:{"Content-Type":"multipart/form-data"},timeout:15000})).data.data;
 },
 async list(){return (await api.get<{data:BetaFeedbackItem[]}>("/beta-feedback")).data.data;},
 async summary(){return (await api.get<{data:{total:number;open:number;p0_open:number;p1_open:number;fixed:number}}>("/beta-feedback/summary")).data.data;},
 async update(id:string,status:string,resolution_notes?:string){return (await api.patch(`/beta-feedback/${id}`,{status,resolution_notes:resolution_notes??null})).data.data;},
 async attachment(id:string){return (await api.get(`/beta-feedback/${id}/attachment`,{responseType:"blob"})).data as Blob;},
};
