export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  // Solo devuelve el API key para que el frontend llame directo
  return res.status(200).json({ 
    key: process.env.CONTIFICO_API_KEY || '',
    base: 'https://api.contifico.com/sistema/api/v1'
  });
}
