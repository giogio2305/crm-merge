const { defineConfig } = require('cypress')
const fs = require('fs')
const xlsx= require('xlsx')

module.exports = defineConfig({
  e2e: {
    setupNodeEvents(on, config) {
      on('task', {
        readXlsx({ file, sheet, skip, limit }) {
          const buf = fs.readFileSync(file);
          const workbook = xlsx.read(buf, { type: 'buffer' })
          const rows = xlsx.utils.sheet_to_json(workbook.Sheets[sheet])
          return rows.slice(skip, skip + limit)
        }
        
      })
    },
  },
})