/**
 * Funções Utilitárias para o Projeto
 */
function getTimestamp() {
  return new Date().toISOString();
}

function formatCurrency(value) {
  return "R$ " + parseFloat(value).toFixed(2).replace('.', ',');
}

function generateId() {
  return Utilities.getUuid();
}

function getUserSession() {
  // Simula obtenção do usuário atual. Em GAS, pode-se usar Session.getActiveUser().getEmail()
  return {
    userId: 'USER_123',
    email: Session.getActiveUser().getEmail()
  };
}
