import { api } from '@/lib/api'; import type {ApiResponse} from '@/types/finance'; import type {ForecastRun,PlanningVersion,Artifact} from '@/types/advanced-forecasting';
export const advancedFinanceService={
 async runForecast(payload:Record<string,unknown>):Promise<ForecastRun>{return (await api.post<ApiResponse<ForecastRun>>('/advanced-forecast/run',payload)).data.data},
 async powerOfOne(payload:Record<string,unknown>){return (await api.post<ApiResponse<any>>('/advanced-forecast/power-of-one',payload)).data.data},
 async versions():Promise<PlanningVersion[]>{return (await api.get<ApiResponse<PlanningVersion[]>>('/native-planning/versions')).data.data},
 async createVersion(payload:Record<string,unknown>):Promise<PlanningVersion>{return (await api.post<ApiResponse<PlanningVersion>>('/native-planning/versions',payload)).data.data},
 async getVersion(id:string):Promise<PlanningVersion>{return (await api.get<ApiResponse<PlanningVersion>>(`/native-planning/versions/${id}`)).data.data},
 async saveLines(id:string,lines:Array<Record<string,unknown>>):Promise<PlanningVersion>{return (await api.put<ApiResponse<PlanningVersion>>(`/native-planning/versions/${id}/lines`,lines)).data.data},
 async generateBoardPack(payload:Record<string,unknown>):Promise<Artifact[]>{return (await api.post<ApiResponse<Artifact[]>>('/board-packs/generate',payload,{timeout:180000})).data.data},
 async artifacts():Promise<Artifact[]>{return (await api.get<ApiResponse<Artifact[]>>('/board-packs/artifacts')).data.data},
};
