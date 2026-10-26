// postcss.config.js
console.log("✔ Tailwind plugin loaded")

import tailwindcss from '@tailwindcss/postcss'
import autoprefixer from 'autoprefixer'

export default {
  plugins: [tailwindcss(), autoprefixer()]
}