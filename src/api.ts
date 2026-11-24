import type { ApiResponse, Session } from "./types";
const API_URL = import.meta.env.VITE_GAS_WEB_APP_URL as string;
export async function api<T>(action:string, payload:Record<string,unknown> = {}):Promise<T> {
  if (!API_URL) throw new Error("أضف رابط Apps Script في VITE_GAS_WEB_APP_URL");
  const response = await fetch(API_URL, {
    method:"POST",
    headers:{"Content-Type":"text/plain;charset=utf-8"},
    body:JSON.stringify({action,...payload}),
    redirect:"follow"
  });
  const result = await response.json() as ApiResponse<T>;
  if (!result.ok) throw new Error(result.error || "حدث خطأ غير معروف");
  return (result.data ?? result) as T;
}
export async function login(username:string,password:string):Promise<Session> {
  const response = await fetch(API_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({action:"login",username,password}),redirect:"follow"});
  const result = await response.json();
  if (!result.ok) throw new Error(result.error || "تعذر تسجيل الدخول");
  return result as Session;
}