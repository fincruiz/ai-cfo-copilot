import { api } from "@/lib/api";
export type DemoAnswer={answer:string;mode:string;evidence:Array<{label:string;value:string;source:string}>;confidence:"high"|"medium"|"low";confidence_reason:string;suggested_questions:string[];visualization?:{type:string;title:string;subtitle?:string;labels:string[];series:Array<{name:string;data:number[]}>;value_format?:string;currency?:string}|null;action?:{label:string;demo_anchor?:string}|null};
export const demoService={async ask(question:string){return (await api.post<{data:DemoAnswer}>("/demo/ask",{question},{timeout:22000})).data.data;}};
