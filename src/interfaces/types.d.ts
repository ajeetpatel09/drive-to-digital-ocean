export interface CustomResponse {
  data: any;
  error: any;
  message: string;
}

export interface JwtPayload {
  id: number;
  isAdmin: boolean;
  collegeId?: number;
}

export type RequestContext = {
  auth?: JwtPayload;
  tenantCollegeId?: number;
};
