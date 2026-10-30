import {randomBytes,createHmac,timingSafeEqual} from 'node:crypto';

export function createSessionManager(){
  const secret=randomBytes(32),cookieName='slidehandoff-session';
  const sign=value=>createHmac('sha256',secret).update(value).digest('hex');
  const equal=(a,b)=>typeof a==='string'&&/^[a-f0-9]{64}$/.test(a)&&typeof b==='string'&&a.length===b.length&&timingSafeEqual(Buffer.from(a),Buffer.from(b));
  function read(req){
    const matches=(req.headers.cookie??'').split(';').map(s=>s.trim()).filter(s=>s.startsWith(cookieName+'='));
    if(matches.length!==1)return null;
    const token=matches[0].slice(cookieName.length+1);
    if(!/^[a-f0-9]{64}\.[a-f0-9]{64}$/.test(token))return null;
    const [ownerId,signature]=token.split('.');
    return equal(signature,sign(ownerId))?{ownerId,csrfToken:sign('csrf:'+token)}:null;
  }
  return {
    read,
    bootstrap(req,res){
      const existing=read(req);if(existing)return existing;
      const ownerId=randomBytes(32).toString('hex'),token=ownerId+'.'+sign(ownerId);
      res.setHeader('set-cookie',`${cookieName}=${token}; HttpOnly; SameSite=Strict; Path=/`);
      return {ownerId,csrfToken:sign('csrf:'+token)};
    },
    permitsWrite(req,session){return Boolean(session&&equal(req.headers['x-csrf-token'],session.csrfToken));}
  };
}
