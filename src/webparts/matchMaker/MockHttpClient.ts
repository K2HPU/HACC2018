//Created file to use with getList method
import { ISPList } from './MatchMakerWebPart';  
  
export default class MockHttpClient {  
    private static _items: ISPList[] = [{ Title: 'FileName', ResourceType: 'LessonPlan', SubjectArea: 'Math', TargetAudience: '1stGrade' },];  
    public static get(restUrl: string, options?: any): Promise<ISPList[]> {  
      return new Promise<ISPList[]>((resolve) => {  
            resolve(MockHttpClient._items);  
        });  
    }  
}  