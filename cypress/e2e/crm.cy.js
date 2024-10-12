
describe('CRM ID Verification Automation', () => {
  let crmData = [];

  before(() => {
    // // Step 1: Load CRM Data from Excel
    // const workbook = XLSX.readFile('../fixtures/crm_data.xlsx');
    // const sheet_name_list = workbook.SheetNames;
    // const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    // crmData = data; // CRM data loaded
    // Step 2: Navigate to website
          cy.visit('https://together.american.edu/index.html#supporters');
          cy.get('#enLoginUsername').type("nd6003a@american.edu");
          cy.get('#enLoginPassword').type("@Duchain0205376003");
          cy.get('.button--login').click();
  });

  context('xlsx file', () => {
    it('Read excel file', () => {
      cy.get('[href="#reports"]').click();
      cy.contains('Lookup supporter').click();
      cy.get('.btn-advanced').click();
      cy.task('readXlsx', { file: "D:/engaging/cypress/fixtures/crm_data.xlsx", sheet: "Sheet1",skip: 255, limit: 5 })
      .then((rows) => {
        rows.forEach((row) => {
          const crmId = row['crm_id'];
          const firstName = row['first_name'];
          const lastName = row['last_name'];
          const emailAddress = row['email_address'];
          // Do something with the CRM ID
          cy.log(`CRM ID: ${crmId}`);
          cy.get(':nth-child(1) > :nth-child(3) > input').clear();
          cy.get(':nth-child(1) > :nth-child(3) > input').as('inputField');
          cy.get('@inputField').type(crmId);
          cy.get('.buttons--left > .button').click();
          cy.wait(2000);

          cy.get('.tableContainer').find('tr').then((tableRows) => {
            if (tableRows.length >= 3) {
              cy.get('.tableContainer').find('tr').each(($row) => {
                cy.contains(emailAddress).should('be.visible');
              });
              cy.get('thead > tr > :nth-child(1) > input').check();
              cy.contains('Merge').click();
              cy.get('thead > tr > :nth-child(5) > .button').click();
              cy.get(".button.btn-save").click();
              cy.get('.message__actions__confirm').click();
              cy.wait(2000);
              // check if the table has less than 3 rows after merging
              cy.get('.tableContainer').find('tr').then((tableRows) => {
                if (tableRows.length < 3) {
                  cy.log(`CRM ID ${crmId} successfully merged`);
                } else {
                  cy.log(`CRM ID ${crmId} failed to merge`);
                }
              })
            } else {
              // Skip checks if table has less than 3 rows
              cy.log(`Skipping checks for CRM ID ${crmId} due to no merge conflict`);
            }
          });



        })
      })
    })
  })

  after(() => {
    // Step 8: Save Final CRM List
    // const newWorkbook = XLSX.utils.book_new();
    // const newSheet = XLSX.utils.json_to_sheet(crmData);
    // XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'FilteredCRM');
    // XLSX.writeFile(newWorkbook, 'path/to/output/filtered_crm_data.xlsx');
  });
});