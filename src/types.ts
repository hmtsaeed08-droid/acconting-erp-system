export type Role = "SUPER_ADMIN" | "ADMIN_1" | "ADMIN_2" | "VIEWER";
export interface User { id:string; name:string; employeeNo:string; username:string; role:Role; }
export interface ApiResponse<T> { ok:boolean; data?:T; error?:string; }
export interface Session { token:string; expiresAt:string; mustChangePassword:boolean; user:User; }