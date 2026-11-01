export async function createTestSession(base){
  const response=await fetch(base+'/api/session');
  if(!response.ok)throw new Error('Session bootstrap failed');
  const cookie=response.headers.get('set-cookie').split(';')[0];
  const {csrfToken}=await response.json();
  return (route,options={})=>fetch(new URL(route,base),{...options,headers:{cookie,'x-csrf-token':csrfToken,...options.headers}});
}
