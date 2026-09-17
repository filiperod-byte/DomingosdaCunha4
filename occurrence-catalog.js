/* Catálogo partilhado entre interface e backend. */
const DC4_OCCURRENCE_CATALOG = {
  floors: [10,9,8,7,6,5,4,3,2,0,-1,-2,-3],
  statuses: {RECEBIDO:'Recebido',EM_ANALISE:'Em análise',AGENDADO:'Intervenção agendada',RESOLVIDO:'Resolvido',REJEITADO:'Não validado'},
  categories: [
    {id:'iluminacao',label:'Iluminação',scope:'floor',reasons:['Luz apagada','Luz intermitente','Sensor / temporizador','Outro']},
    {id:'portas',label:'Portas e portas corta-fogo',scope:'floor',reasons:['Não fecha corretamente','Fechadura / puxador','Porta danificada','Acesso obstruído','Outro']},
    {id:'limpeza',label:'Limpeza',scope:'both',reasons:['Sujidade','Resíduos abandonados','Derrame','Outro']},
    {id:'conservacao',label:'Pavimento, paredes e teto',scope:'floor',reasons:['Dano / peça solta','Infiltração / humidade','Outro']},
    {id:'elevadores',label:'Elevadores',scope:'general',reasons:['Fora de serviço','Portas não abrem / fecham','Ruído / funcionamento irregular','Outro']},
    {id:'portao',label:'Portão da garagem',scope:'general',reasons:['Não abre / fecha','Sensor / comando','Ruído / dano','Outro']},
    {id:'entrada',label:'Porta de entrada / intercomunicador',scope:'general',reasons:['Não abre / fecha','Código / chave / fechadura','Intercomunicador','Outro']},
    {id:'agua',label:'Abastecimento de água / bombas',scope:'general',reasons:['Falta de água','Pressão irregular','Fuga de água','Ruído nas bombas','Outro']},
    {id:'exterior',label:'Exterior / fachada',scope:'general',reasons:['Dano / peça solta','Infiltração','Outro']},
    {id:'outro',label:'Outro problema',scope:'both',reasons:['Outro']}
  ]
};
if (typeof window !== 'undefined') window.DC4_OCCURRENCE_CATALOG = DC4_OCCURRENCE_CATALOG;
