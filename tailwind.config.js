/** Tailwind の設定。使うクラスだけを先に作るため、原本を content に並べる。
 *  cdn.tailwindcss.com（ブラウザ内で CSS を生成する版）は使わない。
 *  以前は App.html の中に tailwind.config = {…} として書いてあった。 */
module.exports = {
  content: ["./src/**/*.jsx", "./App.html"],
  theme: {
    extend: {
      colors: {
        brand: { 50: '#f0f9ff', 100: '#e0f2fe', 500: '#0ea5e9', 600: '#0284c7', 900: '#0c4a6e' },
        accent: { 50: '#fff1f2', 100: '#ffe4e6', 500: '#f43f5e', 600: '#e11d48', 900: '#881337' },
        surface: '#f8fafc',
      },
      fontFamily: { 'rounded': ['"M PLUS Rounded 1c"', 'sans-serif'] },
      boxShadow: {
        'soft': '0 4px 20px -2px rgba(0, 0, 0, 0.05)',
        'float': '0 10px 40px -10px rgba(0,0,0,0.12)',
        'inner-soft': 'inset 0 2px 4px 0 rgba(0, 0, 0, 0.04)',
      },
      zIndex: { 'overlay': 8000, 'modal': 9000, 'toast': 9999 }
    }
  }
};
