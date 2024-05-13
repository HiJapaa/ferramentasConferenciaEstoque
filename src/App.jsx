import { useState, useEffect } from 'react'
import { db } from './services/firebaseConnection'
import { addDoc, collection, getDocs, getDoc, doc } from 'firebase/firestore'
import * as XLSX from 'xlsx';

const listRef = collection(db, 'teste')


function App() {
  const [relatorios, setRelatorios] = useState([])
  let lista = []

  useEffect(() => {
    async function loadTexts() {
      const querySnapshot = await getDocs(listRef)
        .then((snapshot) => {
          snapshot.forEach(doc => {
            lista.push({
              id: doc.id,
              data: doc.data().data,
              resp: doc.data().resp,
              categoria: doc.data().categoria,
              loja: doc.data().loja,
              text: doc.data().texto
            })
          })
          if (snapshot.docs.size === 0) {
            console.log('Vazio')
            return
          }
          setRelatorios(lista)
        })
        .catch((err) => {
          console.log('Erro ao ler', err)
        })
    }
    loadTexts()
  }, [])


  // Leitura do arquivo XLSX
  function estoque(e) {
    const file = e.target.files[0];

    const reader = new FileReader();

    reader.onload = (e) => {
      const workbook = XLSX.read(e.target.result, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      processCSVEstoque(data)

      alert('Terminou! Clique no botão "Atualizar conferências"')
    };

    reader.readAsBinaryString(file);
  };

  let dataEstoque = []
  let lucasData = []
  let filintoData = []
  let roiData = []
  let caceresData = []
  let sorrisoData = []
  let estoqueAdmData = []
  let vgShoppingData = []
  let barraData = []
  let coutoData = []
  let primaveraData = []
  let pontesData = []
  let coliderData = []
  let mirassolData = []
  let guarantaData = []
  let jaciaraData = []
  let comodoroData = []
  let altoAraguaiaData = []
  let ariquemesData = []
  let estoqueRoData = []
  let querenciaData = []
  let jaruData = []
  let jiParanaData = []
  let peixotoData = []
  let rolimData = []
  let pimentaData = []
  let vilhenaData = []
  let cacoalData = []
  let setembroData = []
  let jatuaranaData = []
  let joseData = []
  let confresaData = []
  let xavantinaData = []

  function pshAcessorios(data, e) {
    data.push({
      Serial: e[3],
      Quantidade: e[13] * (-1)
    })
  }

  function psh(data, e) {
    data.push({
      Serial: e[12],
      Quantidade: e[18] * (-1)
    })
  }

  // Processamento do arquivo XLSX
  function processCSVEstoque(content) {
    content.forEach(e => {
      if (categoria == 'Acessorios') {
        dataEstoque.push({
          Filial: e[1],
          Serial: e[3],
          Quantidade: e[13] * (-1)
        })

        if (e[1] == 'MS LUCAS') {
          pshAcessorios(lucasData, e)
        }

        if (e[1] == 'MS FILINTO') {
          pshAcessorios(filintoData, e)
        }

        if (e[1] == 'MS ROI LLA') {
          pshAcessorios(roiData, e)
        }

        if (e[1] == 'MS CÁCERES') {
          pshAcessorios(caceresData, e)
        }

        if (e[1] == 'MS SORRISO') {
          pshAcessorios(sorrisoData, e)
        }

        if (e[1] == 'ESTOQUE ADM') {
          pshAcessorios(estoqueAdmData, e)
        }

        if (e[1] == 'MS VG Shopping') {
          pshAcessorios(vgShoppingData, e)
        }

        if (e[1] == 'MS BARRA DO GARÇAS') {
          pshAcessorios(barraData, e)
        }

        if (e[1] == 'MS COUTO') {
          pshAcessorios(coutoData, e)
        }

        if (e[1] == 'MS PRIMAVERA DO LESTE') {
          pshAcessorios(primaveraData, e)
        }

        if (e[1] == 'MS PONTES E LACERDA') {
          pshAcessorios(pontesData, e)
        }

        if (e[1] == 'MS COLIDER') {
          pshAcessorios(coliderData, e)
        }

        if (e[1] == 'MS MIRASSOL') {
          pshAcessorios(mirassolData, e)
        }

        if (e[1] == 'MS GUARANTÃ DO NORTE') {
          pshAcessorios(guarantaData, e)
        }

        if (e[1] == 'MS JACIARA') {
          pshAcessorios(jaciaraData, e)
        }

        if (e[1] == 'MS COMODORO') {
          pshAcessorios(comodoroData, e)
        }

        if (e[1] == 'MS ALTO ARAGUAIA') {
          pshAcessorios(altoAraguaiaData, e)
        }

        if (e[1] == 'MS ARIQUEMES') {
          pshAcessorios(ariquemesData, e)
        }

        if (e[1] == 'ESTOQUE ADM RO') {
          pshAcessorios(estoqueRoData, e)
        }

        if (e[1] == 'MS QUERÊNCIA') {
          pshAcessorios(querenciaData, e)
        }

        if (e[1] == 'MS JARU') {
          pshAcessorios(jaruData, e)
        }

        if (e[1] == 'MS JI-PARANA') {
          pshAcessorios(jiParanaData, e)
        }

        if (e[1] == 'MS PEIXOTO DE AZEVEDO') {
          pshAcessorios(peixotoData, e)
        }

        if (e[1] == 'MS ROLIM DE MOURA') {
          pshAcessorios(rolimData, e)
        }

        if (e[1] == 'MS PIMENTA BUENO') {
          pshAcessorios(pimentaData, e)
        }

        if (e[1] == 'MS VILHENA') {
          pshAcessorios(vilhenaData, e)
        }

        if (e[1] == 'MS CACOAL') {
          pshAcessorios(cacoalData, e)
        }

        if (e[1] == 'MS PV - 07 SETEMBRO') {
          pshAcessorios(setembroData, e)
        }

        if (e[1] == 'MS PV - JATUARANA') {
          pshAcessorios(jatuaranaData, e)
        }

        if (e[1] == 'MS PV - JOSE AMADOR') {
          pshAcessorios(joseData, e)
        }

        if (e[1] == 'MS CONFRESA') {
          pshAcessorios(confresaData, e)
        }

        if (e[1] == 'MS NOVA XAVANTINA') {
          pshAcessorios(xavantinaData, e)
        }
      } else {
        dataEstoque.push({
          Filial: e[1],
          Serial: e[12],
          Quantidade: e[18] * (-1)
        })

        if (e[1] == 'MS LUCAS') {
          psh(lucasData, e)
        }

        if (e[1] == 'MS FILINTO') {
          psh(filintoData, e)
        }

        if (e[1] == 'MS ROI LLA') {
          psh(roiData, e)
        }

        if (e[1] == 'MS CÁCERES') {
          psh(caceresData, e)
        }

        if (e[1] == 'MS SORRISO') {
          psh(sorrisoData, e)
        }

        if (e[1] == 'ESTOQUE ADM') {
          psh(estoqueAdmData, e)
        }

        if (e[1] == 'MS VG Shopping') {
          psh(vgShoppingData, e)
        }

        if (e[1] == 'MS BARRA DO GARÇAS') {
          psh(barraData, e)
        }

        if (e[1] == 'MS COUTO') {
          psh(coutoData, e)
        }

        if (e[1] == 'MS PRIMAVERA DO LESTE') {
          psh(primaveraData, e)
        }

        if (e[1] == 'MS PONTES E LACERDA') {
          psh(pontesData, e)
        }

        if (e[1] == 'MS COLIDER') {
          psh(coliderData, e)
        }

        if (e[1] == 'MS MIRASSOL') {
          psh(mirassolData, e)
        }

        if (e[1] == 'MS GUARANTÃ DO NORTE') {
          psh(guarantaData, e)
        }

        if (e[1] == 'MS JACIARA') {
          psh(jaciaraData, e)
        }

        if (e[1] == 'MS COMODORO') {
          psh(comodoroData, e)
        }

        if (e[1] == 'MS ALTO ARAGUAIA') {
          psh(altoAraguaiaData, e)
        }

        if (e[1] == 'MS ARIQUEMES') {
          psh(ariquemesData, e)
        }

        if (e[1] == 'ESTOQUE ADM RO') {
          psh(estoqueRoData, e)
        }

        if (e[1] == 'MS QUERÊNCIA') {
          psh(querenciaData, e)
        }

        if (e[1] == 'MS JARU') {
          psh(jaruData, e)
        }

        if (e[1] == 'MS JI-PARANA') {
          psh(jiParanaData, e)
        }

        if (e[1] == 'MS PEIXOTO DE AZEVEDO') {
          psh(peixotoData, e)
        }

        if (e[1] == 'MS ROLIM DE MOURA') {
          psh(rolimData, e)
        }

        if (e[1] == 'MS PIMENTA BUENO') {
          psh(pimentaData, e)
        }

        if (e[1] == 'MS VILHENA') {
          psh(vilhenaData, e)
        }

        if (e[1] == 'MS CACOAL') {
          psh(cacoalData, e)
        }

        if (e[1] == 'MS PV - 07 SETEMBRO') {
          psh(setembroData, e)
        }

        if (e[1] == 'MS PV - JATUARANA') {
          psh(jatuaranaData, e)
        }

        if (e[1] == 'MS PV - JOSE AMADOR') {
          psh(joseData, e)
        }

        if (e[1] == 'MS CONFRESA') {
          psh(confresaData, e)
        }

        if (e[1] == 'MS NOVA XAVANTINA') {
          psh(xavantinaData, e)
        }
      }
    })
  }

  let lojas = ['MS LUCAS', 'MS FILINTO', 'MS ROI LLA', 'MS CÁCERES', 'MS SORRISO', 'ESTOQUE ADM', 'MS VG Shopping', 'MS BARRA DO GARÇAS', 'MS COUTO', 'MS PRIMAVERA DO LESTE', 'MS PONTES E LACERDA', 'MS COLIDER', 'MS MIRASSOL', 'MS GUARANTÃ DO NORTE', 'MS JACIARA', 'MS COMODORO', 'MS ALTO ARAGUAIA', 'MS ARIQUEMES', 'ESTOQUE ADM RO', 'MS QUERÊNCIA', 'MS JARU', 'MS JI-PARANA', 'MS PEIXOTO DE AZEVEDO', 'MS ROLIM DE MOURA', 'MS PIMENTA BUENO', 'MS VILHENA', 'MS CACOAL', 'MS PV - 07 SETEMBRO', 'MS PV - JATUARANA', 'MS PV - JOSE AMADOR', 'MS CONFRESA', 'MS NOVA XAVANTINA']
  let [faltantes, setFaltantes] = useState([])
  let feitas = []

  function fizeram() {
    if (categoria == 'Diversos') {
      relatorios.forEach(e => {
        if (e.categoria == 'Diversos') {
          let achou = lojas.find(element => element === e.loja)
          feitas.push(achou)
        }
      })
      console.log(feitas)
      setFaltantes(lojas.filter(a => !feitas.includes(a)))
    } else if (categoria == 'Chips') {
      relatorios.forEach(e => {
        if (e.categoria == 'Chips') {
          let achou = lojas.find(element => element === e.loja)
          feitas.push(achou)
        }
      })
      console.log(feitas)
      setFaltantes(lojas.filter(a => !feitas.includes(a)))
    } else if (categoria == 'ProdutosVivo') {
      relatorios.forEach(e => {
        if (e.categoria == 'ProdutosVivo') {
          let achou = lojas.find(element => element === e.loja)
          feitas.push(achou)
        }
      })
      console.log(feitas)
      setFaltantes(lojas.filter(a => !feitas.includes(a)))
    } else if (categoria == 'Acessorios') {
      relatorios.forEach(e => {
        if (e.categoria == 'Acessorios') {
          let achou = lojas.find(element => element === e.loja)
          feitas.push(achou)
        }
      })
      console.log(feitas)
      setFaltantes(lojas.filter(a => !feitas.includes(a)))
    }

    setListaFaltantes()
  }

  let lucasDataConf = []
  let filintoDataConf = []
  let roiDataConf = []
  let caceresDataConf = []
  let sorrisoDataConf = []
  let estoqueAdmDataConf = []
  let vgShoppingDataConf = []
  let barraDataConf = []
  let coutoDataConf = []
  let primaveraDataConf = []
  let pontesDataConf = []
  let coliderDataConf = []
  let mirassolDataConf = []
  let guarantaDataConf = []
  let jaciaraDataConf = []
  let comodoroDataConf = []
  let altoAraguaiaDataConf = []
  let ariquemesDataConf = []
  let estoqueRoDataConf = []
  let querenciaDataConf = []
  let jaruDataConf = []
  let jiParanaDataConf = []
  let peixotoDataConf = []
  let rolimDataConf = []
  let pimentaDataConf = []
  let vilhenaDataConf = []
  let cacoalDataConf = []
  let setembroDataConf = []
  let jatuaranaDataConf = []
  let joseDataConf = []
  let confresaDataConf = []
  let xavantinaDataConf = []

  let lucasChipsConf = []
  let filintoChipsConf = []
  let roiChipsConf = []
  let caceresChipsConf = []
  let sorrisoChipsConf = []
  let estoqueAdmChipsConf = []
  let vgShoppingChipsConf = []
  let barraChipsConf = []
  let coutoChipsConf = []
  let primaveraChipsConf = []
  let pontesChipsConf = []
  let coliderChipsConf = []
  let mirassolChipsConf = []
  let guarantaChipsConf = []
  let jaciaraChipsConf = []
  let comodoroChipsConf = []
  let altoAraguaiaChipsConf = []
  let ariquemesChipsConf = []
  let estoqueRoChipsConf = []
  let querenciaChipsConf = []
  let jaruChipsConf = []
  let jiParanaChipsConf = []
  let peixotoChipsConf = []
  let rolimChipsConf = []
  let pimentaChipsConf = []
  let vilhenaChipsConf = []
  let cacoalChipsConf = []
  let setembroChipsConf = []
  let jatuaranaChipsConf = []
  let joseChipsConf = []
  let confresaChipsConf = []
  let xavantinaChipsConf = []

  let lucasProdutosVivoConf = []
  let filintoProdutosVivoConf = []
  let roiProdutosVivoConf = []
  let caceresProdutosVivoConf = []
  let sorrisoProdutosVivoConf = []
  let estoqueAdmProdutosVivoConf = []
  let vgShoppingProdutosVivoConf = []
  let barraProdutosVivoConf = []
  let coutoProdutosVivoConf = []
  let primaveraProdutosVivoConf = []
  let pontesProdutosVivoConf = []
  let coliderProdutosVivoConf = []
  let mirassolProdutosVivoConf = []
  let guarantaProdutosVivoConf = []
  let jaciaraProdutosVivoConf = []
  let comodoroProdutosVivoConf = []
  let altoAraguaiaProdutosVivoConf = []
  let ariquemesProdutosVivoConf = []
  let estoqueRoProdutosVivoConf = []
  let querenciaProdutosVivoConf = []
  let jaruProdutosVivoConf = []
  let jiParanaProdutosVivoConf = []
  let peixotoProdutosVivoConf = []
  let rolimProdutosVivoConf = []
  let pimentaProdutosVivoConf = []
  let vilhenaProdutosVivoConf = []
  let cacoalProdutosVivoConf = []
  let setembroProdutosVivoConf = []
  let jatuaranaProdutosVivoConf = []
  let joseProdutosVivoConf = []
  let confresaProdutosVivoConf = []
  let xavantinaProdutosVivoConf = []

  let lucasAcessoriosConf = []
  let filintoAcessoriosConf = []
  let roiAcessoriosConf = []
  let caceresAcessoriosConf = []
  let sorrisoAcessoriosConf = []
  let estoqueAdmAcessoriosConf = []
  let vgShoppingAcessoriosConf = []
  let barraAcessoriosConf = []
  let coutoAcessoriosConf = []
  let primaveraAcessoriosConf = []
  let pontesAcessoriosConf = []
  let coliderAcessoriosConf = []
  let mirassolAcessoriosConf = []
  let guarantaAcessoriosConf = []
  let jaciaraAcessoriosConf = []
  let comodoroAcessoriosConf = []
  let altoAraguaiaAcessoriosConf = []
  let ariquemesAcessoriosConf = []
  let estoqueRoAcessoriosConf = []
  let querenciaAcessoriosConf = []
  let jaruAcessoriosConf = []
  let jiParanaAcessoriosConf = []
  let peixotoAcessoriosConf = []
  let rolimAcessoriosConf = []
  let pimentaAcessoriosConf = []
  let vilhenaAcessoriosConf = []
  let cacoalAcessoriosConf = []
  let setembroAcessoriosConf = []
  let jatuaranaAcessoriosConf = []
  let joseAcessoriosConf = []
  let confresaAcessoriosConf = []
  let xavantinaAcessoriosConf = []

  function pshConf(dataConf, chipsConf, produtosVivoConf, acessoriosConf, e) {
    if (e.categoria == 'Diversos') {
      dataConf.push(e.text)
    } else if (e.categoria == 'Chips') {
      chipsConf.push(e.text)
    } else if (e.categoria == 'ProdutosVivo') {
      produtosVivoConf.push(e.text)
    } else if (e.categoria == 'Acessorios') {
      acessoriosConf.push(e.text)
    }
  }

  function conferencias() {
    relatorios.forEach(e => {
      if (e.loja == 'MS LUCAS') {
        pshConf(lucasDataConf, lucasChipsConf, lucasProdutosVivoConf, lucasAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   lucasDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   lucasChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   lucasProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   lucasAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS FILINTO') {
        pshConf(filintoDataConf, filintoChipsConf, filintoProdutosVivoConf, filintoAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   filintoDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   filintoChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   filintoProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   filintoAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS ROI LLA') {
        pshConf(roiDataConf, roiChipsConf, roiProdutosVivoConf, roiAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   roiDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   roiChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   roiProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   roiAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS CÁCERES') {
        pshConf(caceresDataConf, caceresChipsConf, caceresProdutosVivoConf, caceresAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   caceresDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   caceresChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   caceresProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   caceresAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS SORRISO') {
        pshConf(sorrisoDataConf, sorrisoChipsConf, sorrisoProdutosVivoConf, sorrisoAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   sorrisoDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   sorrisoChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   sorrisoProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   sorrisoAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'ESTOQUE ADM') {
        pshConf(estoqueAdmDataConf, estoqueAdmChipsConf, estoqueAdmProdutosVivoConf, estoqueAdmAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   estoqueAdmDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   estoqueAdmChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   estoqueAdmProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   estoqueAdmAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS VG Shopping') {
        pshConf(vgShoppingDataConf, vgShoppingChipsConf, vgShoppingProdutosVivoConf, vgShoppingAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   vgShoppingDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   vgShoppingChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   vgShoppingProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   vgShoppingAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS BARRA DO GARÇAS') {
        pshConf(barraDataConf, barraChipsConf, barraProdutosVivoConf, barraAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   barraDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   barraChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   barraProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   barraAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS COUTO') {
        pshConf(coutoDataConf, coutoChipsConf, coutoProdutosVivoConf, coutoAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   coutoDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   coutoChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   coutoProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   coutoAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS PRIMAVERA DO LESTE') {
        pshConf(primaveraDataConf, primaveraChipsConf, primaveraProdutosVivoConf, primaveraAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   primaveraDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   primaveraChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   primaveraProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   primaveraAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS PONTES E LACERDA') {
        pshConf(pontesDataConf, pontesChipsConf, pontesProdutosVivoConf, pontesAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   pontesDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   pontesChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   pontesProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   pontesAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS COLIDER') {
        pshConf(coliderDataConf, coliderChipsConf, coliderProdutosVivoConf, coliderAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   coliderDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   coliderChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   coliderProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   coliderAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS MIRASSOL') {
        pshConf(mirassolDataConf, mirassolChipsConf, mirassolProdutosVivoConf, mirassolAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   mirassolDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   mirassolChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   mirassolProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   mirassolAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS GUARANTÃ DO NORTE') {
        pshConf(guarantaDataConf, guarantaChipsConf, guarantaProdutosVivoConf, guarantaAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   guarantaDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   guarantaChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   guarantaProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   guarantaAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS JACIARA') {
        pshConf(jaciaraDataConf, jaciaraChipsConf, jaciaraProdutosVivoConf, jaciaraAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   jaciaraDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   jaciaraChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   jaciaraProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   jaciaraAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS COMODORO') {
        pshConf(comodoroDataConf, comodoroChipsConf, comodoroProdutosVivoConf, comodoroAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   comodoroDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   comodoroChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   comodoroProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   comodoroAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS ALTO ARAGUAIA') {
        pshConf(altoAraguaiaDataConf, altoAraguaiaChipsConf, altoAraguaiaProdutosVivoConf, altoAraguaiaAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   altoAraguaiaDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   altoAraguaiaChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   altoAraguaiaProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   altoAraguaiaAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS ARIQUEMES') {
        pshConf(ariquemesDataConf, ariquemesChipsConf, ariquemesProdutosVivoConf, ariquemesAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   ariquemesDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   ariquemesChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   ariquemesProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   ariquemesAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'ESTOQUE ADM RO') {
        pshConf(estoqueRoDataConf, estoqueRoChipsConf, estoqueRoProdutosVivoConf, estoqueRoAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   estoqueRoDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   estoqueRoChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   estoqueRoProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   estoqueRoAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS QUERÊNCIA') {
        pshConf(querenciaDataConf, querenciaChipsConf, querenciaProdutosVivoConf, querenciaAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   querenciaDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   querenciaChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   querenciaProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   querenciaAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS JARU') {
        pshConf(jaruDataConf, jaruChipsConf, jaruProdutosVivoConf, jaruAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   jaruDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   jaruChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   jaruProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   jaruAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS JI-PARANA') {
        pshConf(jiParanaDataConf, jiParanaChipsConf, jiParanaProdutosVivoConf, jiParanaAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   jiParanaDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   jiParanaChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   jiParanaProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   jiParanaAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS PEIXOTO DE AZEVEDO') {
        pshConf(peixotoDataConf, peixotoChipsConf, peixotoProdutosVivoConf, peixotoAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   peixotoDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   peixotoChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   peixotoProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   peixotoAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS ROLIM DE MOURA') {
        pshConf(rolimDataConf, rolimChipsConf, rolimProdutosVivoConf, rolimAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   rolimDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   rolimChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   rolimProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   rolimAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS PIMENTA BUENO') {
        pshConf(pimentaDataConf, pimentaChipsConf, pimentaProdutosVivoConf, pimentaAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   pimentaDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   pimentaChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   pimentaProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   pimentaAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS VILHENA') {
        pshConf(vilhenaDataConf, vilhenaChipsConf, vilhenaProdutosVivoConf, vilhenaAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   vilhenaDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   vilhenaChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   vilhenaProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   vilhenaAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS CACOAL') {
        pshConf(cacoalDataConf, cacoalChipsConf, cacoalProdutosVivoConf, cacoalAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   cacoalDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   cacoalChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   cacoalProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   cacoalAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS PV - 07 SETEMBRO') {
        pshConf(setembroDataConf, setembroChipsConf, setembroProdutosVivoConf, setembroAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   setembroDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   setembroChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   setembroProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   setembroAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS PV - JATUARANA') {
        pshConf(jatuaranaDataConf, jatuaranaChipsConf, jatuaranaProdutosVivoConf, jatuaranaAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   jatuaranaDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   jatuaranaChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   jatuaranaProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   jatuaranaAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS PV - JOSE AMADOR') {
        pshConf(joseDataConf, joseChipsConf, joseProdutosVivoConf, joseAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   joseDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   joseChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   joseProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   joseAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS CONFRESA') {
        pshConf(confresaDataConf, confresaChipsConf, confresaProdutosVivoConf, confresaAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   confresaDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   confresaChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   confresaProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   confresaAcessoriosConf.push(e.text)
        // }
      }

      if (e.loja == 'MS NOVA XAVANTINA') {
        pshConf(xavantinaDataConf, xavantinaChipsConf, xavantinaProdutosVivoConf, xavantinaAcessoriosConf, e)
        // if (e.categoria == 'Diversos') {
        //   xavantinaDataConf.push(e.text)
        // } else if (e.categoria == 'Chips') {
        //   xavantinaChipsConf.push(e.text)
        // } else if (e.categoria == 'ProdutosVivo') {
        //   xavantinaProdutosVivoConf.push(e.text)
        // } else if (e.categoria == 'Acessorios') {
        //   xavantinaAcessoriosConf.push(e.text)
        // }
      }
    })

    alert('Terminou, clique no botão "Verificar"')
  }

  let lucasDiferenca = []
  let filintoDiferenca = []
  let roiDiferenca = []
  let caceresDiferenca = []
  let sorrisoDiferenca = []
  let estoqueAdmDiferenca = []
  let vgShoppingDiferenca = []
  let barraDiferenca = []
  let coutoDiferenca = []
  let primaveraDiferenca = []
  let pontesDiferenca = []
  let coliderDiferenca = []
  let mirassolDiferenca = []
  let guarantaDiferenca = []
  let jaciaraDiferenca = []
  let comodoroDiferenca = []
  let altoAraguaiaDiferenca = []
  let ariquemesDiferenca = []
  let estoqueRoDiferenca = []
  let querenciaDiferenca = []
  let jaruDiferenca = []
  let jiParanaDiferenca = []
  let peixotoDiferenca = []
  let rolimDiferenca = []
  let pimentaDiferenca = []
  let vilhenaDiferenca = []
  let cacoalDiferenca = []
  let setembroDiferenca = []
  let jatuaranaDiferenca = []
  let joseDiferenca = []
  let confresaDiferenca = []
  let xavantinaDiferenca = []

  function procurar(conf, diff, data) {
    if (conf[0] !== undefined) {
      conf[0].forEach(e => {
        if (e !== '') {
          let index = data.findIndex(element => element.Serial.includes(e))
          if (index < 0) {
            diff.push({
              Serial: e,
              Quantidade: 1
            })
          } else {
            data[index].Quantidade++
          }
        }
      })
    }
    data.forEach(e => {
      if (e.Quantidade !== 0) {
        diff.push(e)
      }
    })
  }


  function procurarAcessorios(conf, diff, data) {
    if (conf[0] !== undefined) {
      conf[0].forEach(e => {
        if (e !== '') {
          let index = data.findIndex(element => element.Serial == e)
          if (index < 0) {
            diff.push({
              Serial: e,
              Quantidade: 1
            })
          } else {
            data[index].Quantidade++
          }
        }
      })
    }
    data.forEach(e => {
      if (e.Quantidade !== 0) {
        diff.push(e)
      }
    })
  }

  function verificarDiff(diff, nome, tempo) {
    setTimeout(() => {
      if (diff.length > 0) {
        exportToXLSX(diff, `${nome}`);
      }
    }, tempo)
  }

  async function leitura() {
    // Quando Diversos estiver marcado
    if (categoria == 'Diversos') {
      procurar(lucasDataConf, lucasDiferenca, lucasData)

      procurar(filintoDataConf, filintoDiferenca, filintoData)

      procurar(roiDataConf, roiDiferenca, roiData)

      procurar(caceresDataConf, caceresDiferenca, caceresData)

      procurar(sorrisoDataConf, sorrisoDiferenca, sorrisoData)

      procurar(estoqueAdmDataConf, estoqueAdmDiferenca, estoqueAdmData)

      procurar(vgShoppingDataConf, vgShoppingDiferenca, vgShoppingData)

      procurar(barraDataConf, barraDiferenca, barraData)

      procurar(coutoDataConf, coutoDiferenca, coutoData)

      procurar(primaveraDataConf, primaveraDiferenca, primaveraData)

      procurar(pontesDataConf, pontesDiferenca, pontesData)

      procurar(coliderDataConf, coliderDiferenca, coliderData)

      procurar(mirassolDataConf, mirassolDiferenca, mirassolData)

      procurar(guarantaDataConf, guarantaDiferenca, guarantaData)

      procurar(jaciaraDataConf, jaciaraDiferenca, jaciaraData)

      procurar(comodoroDataConf, comodoroDiferenca, comodoroData)

      procurar(altoAraguaiaDataConf, altoAraguaiaDiferenca, altoAraguaiaData)

      procurar(ariquemesDataConf, ariquemesDiferenca, ariquemesData)

      procurar(estoqueRoDataConf, estoqueRoDiferenca, estoqueRoData)

      procurar(querenciaDataConf, querenciaDiferenca, querenciaData)

      procurar(jaruDataConf, jaruDiferenca, jaruData)

      procurar(jiParanaDataConf, jiParanaDiferenca, jiParanaData)

      procurar(peixotoDataConf, peixotoDiferenca, peixotoData)

      procurar(rolimDataConf, rolimDiferenca, rolimData)

      procurar(pimentaDataConf, pimentaDiferenca, pimentaData)

      procurar(vilhenaDataConf, vilhenaDiferenca, vilhenaData)

      procurar(cacoalDataConf, cacoalDiferenca, cacoalData)

      procurar(setembroDataConf, setembroDiferenca, setembroData)

      procurar(jatuaranaDataConf, jatuaranaDiferenca, jatuaranaData)

      procurar(joseDataConf, joseDiferenca, joseData)

      procurar(confresaDataConf, confresaDiferenca, confresaData)

      procurar(xavantinaDataConf, xavantinaDiferenca, xavantinaData)


      // Quando CHIP estiver marcado
    } else if (categoria == 'Chips') {
      procurar(lucasChipsConf, lucasDiferenca, lucasData)

      procurar(filintoChipsConf, filintoDiferenca, filintoData)

      procurar(roiChipsConf, roiDiferenca, roiData)

      procurar(caceresChipsConf, caceresDiferenca, caceresData)

      procurar(sorrisoChipsConf, sorrisoDiferenca, sorrisoData)

      procurar(estoqueAdmChipsConf, estoqueAdmDiferenca, estoqueAdmData)

      procurar(vgShoppingChipsConf, vgShoppingDiferenca, vgShoppingData)

      procurar(barraChipsConf, barraDiferenca, barraData)

      procurar(coutoChipsConf, coutoDiferenca, coutoData)

      procurar(primaveraChipsConf, primaveraDiferenca, primaveraData)

      procurar(pontesChipsConf, pontesDiferenca, pontesData)

      procurar(coliderChipsConf, coliderDiferenca, coliderData)

      procurar(mirassolChipsConf, mirassolDiferenca, mirassolData)

      procurar(guarantaChipsConf, guarantaDiferenca, guarantaData)

      procurar(jaciaraChipsConf, jaciaraDiferenca, jaciaraData)

      procurar(comodoroChipsConf, comodoroDiferenca, comodoroData)

      procurar(altoAraguaiaChipsConf, altoAraguaiaDiferenca, altoAraguaiaData)

      procurar(ariquemesChipsConf, ariquemesDiferenca, ariquemesData)

      procurar(estoqueRoChipsConf, estoqueRoDiferenca, estoqueRoData)

      procurar(querenciaChipsConf, querenciaDiferenca, querenciaData)

      procurar(jaruChipsConf, jaruDiferenca, jaruData)

      procurar(jiParanaChipsConf, jiParanaDiferenca, jiParanaData)

      procurar(peixotoChipsConf, peixotoDiferenca, peixotoData)

      procurar(rolimChipsConf, rolimDiferenca, rolimData)

      procurar(pimentaChipsConf, pimentaDiferenca, pimentaData)

      procurar(vilhenaChipsConf, vilhenaDiferenca, vilhenaData)

      procurar(cacoalChipsConf, cacoalDiferenca, cacoalData)

      procurar(setembroChipsConf, setembroDiferenca, setembroData)

      procurar(jatuaranaChipsConf, jatuaranaDiferenca, jatuaranaData)

      procurar(joseChipsConf, joseDiferenca, joseData)

      procurar(confresaChipsConf, confresaDiferenca, confresaData)

      procurar(xavantinaChipsConf, xavantinaDiferenca, xavantinaData)


      // Quando PRODUTOSVIVO estiver marcado
    } else if (categoria == 'ProdutosVivo') {
      procurar(lucasProdutosVivoConf, lucasDiferenca, lucasData)

      procurar(filintoProdutosVivoConf, filintoDiferenca, filintoData)

      procurar(roiProdutosVivoConf, roiDiferenca, roiData)

      procurar(caceresProdutosVivoConf, caceresDiferenca, caceresData)

      procurar(sorrisoProdutosVivoConf, sorrisoDiferenca, sorrisoData)

      procurar(estoqueAdmProdutosVivoConf, estoqueAdmDiferenca, estoqueAdmData)

      procurar(vgShoppingProdutosVivoConf, vgShoppingDiferenca, vgShoppingData)

      procurar(barraProdutosVivoConf, barraDiferenca, barraData)

      procurar(coutoProdutosVivoConf, coutoDiferenca, coutoData)

      procurar(primaveraProdutosVivoConf, primaveraDiferenca, primaveraData)

      procurar(pontesProdutosVivoConf, pontesDiferenca, pontesData)

      procurar(coliderProdutosVivoConf, coliderDiferenca, coliderData)

      procurar(mirassolProdutosVivoConf, mirassolDiferenca, mirassolData)

      procurar(guarantaProdutosVivoConf, guarantaDiferenca, guarantaData)

      procurar(jaciaraProdutosVivoConf, jaciaraDiferenca, jaciaraData)

      procurar(comodoroProdutosVivoConf, comodoroDiferenca, comodoroData)

      procurar(altoAraguaiaProdutosVivoConf, altoAraguaiaDiferenca, altoAraguaiaData)

      procurar(ariquemesProdutosVivoConf, ariquemesDiferenca, ariquemesData)

      procurar(estoqueRoProdutosVivoConf, estoqueRoDiferenca, estoqueRoData)

      procurar(querenciaProdutosVivoConf, querenciaDiferenca, querenciaData)

      procurar(jaruProdutosVivoConf, jaruDiferenca, jaruData)

      procurar(jiParanaProdutosVivoConf, jiParanaDiferenca, jiParanaData)

      procurar(peixotoProdutosVivoConf, peixotoDiferenca, peixotoData)

      procurar(rolimProdutosVivoConf, rolimDiferenca, rolimData)

      procurar(pimentaProdutosVivoConf, pimentaDiferenca, pimentaData)

      procurar(vilhenaProdutosVivoConf, vilhenaDiferenca, vilhenaData)

      procurar(cacoalProdutosVivoConf, cacoalDiferenca, cacoalData)

      procurar(setembroProdutosVivoConf, setembroDiferenca, setembroData)

      procurar(jatuaranaProdutosVivoConf, jatuaranaDiferenca, jatuaranaData)

      procurar(joseProdutosVivoConf, joseDiferenca, joseData)

      procurar(confresaProdutosVivoConf, confresaDiferenca, confresaData)

      procurar(xavantinaProdutosVivoConf, xavantinaDiferenca, xavantinaData)


      // Quando ACESSORIOS estiver marcado
    } else if (categoria == 'Acessorios') {
      procurarAcessorios(lucasAcessoriosConf, lucasDiferenca, lucasData)

      procurarAcessorios(filintoAcessoriosConf, filintoDiferenca, filintoData)

      procurarAcessorios(roiAcessoriosConf, roiDiferenca, roiData)

      procurarAcessorios(caceresAcessoriosConf, caceresDiferenca, caceresData)

      procurarAcessorios(sorrisoAcessoriosConf, sorrisoDiferenca, sorrisoData)

      procurarAcessorios(estoqueAdmAcessoriosConf, estoqueAdmDiferenca, estoqueAdmData)

      procurarAcessorios(vgShoppingAcessoriosConf, vgShoppingDiferenca, vgShoppingData)

      procurarAcessorios(barraAcessoriosConf, barraDiferenca, barraData)

      procurarAcessorios(coutoAcessoriosConf, coutoDiferenca, coutoData)

      procurarAcessorios(primaveraAcessoriosConf, primaveraDiferenca, primaveraData)

      procurarAcessorios(pontesAcessoriosConf, pontesDiferenca, pontesData)

      procurarAcessorios(coliderAcessoriosConf, coliderDiferenca, coliderData)

      procurarAcessorios(mirassolAcessoriosConf, mirassolDiferenca, mirassolData)

      procurarAcessorios(guarantaAcessoriosConf, guarantaDiferenca, guarantaData)

      procurarAcessorios(jaciaraAcessoriosConf, jaciaraDiferenca, jaciaraData)

      procurarAcessorios(comodoroAcessoriosConf, comodoroDiferenca, comodoroData)

      procurarAcessorios(altoAraguaiaAcessoriosConf, altoAraguaiaDiferenca, altoAraguaiaData)

      procurarAcessorios(ariquemesAcessoriosConf, ariquemesDiferenca, ariquemesData)

      procurarAcessorios(estoqueRoAcessoriosConf, estoqueRoDiferenca, estoqueRoData)

      procurarAcessorios(querenciaAcessoriosConf, querenciaDiferenca, querenciaData)

      procurarAcessorios(jaruAcessoriosConf, jaruDiferenca, jaruData)

      procurarAcessorios(jiParanaAcessoriosConf, jiParanaDiferenca, jiParanaData)

      procurarAcessorios(peixotoAcessoriosConf, peixotoDiferenca, peixotoData)

      procurarAcessorios(rolimAcessoriosConf, rolimDiferenca, rolimData)

      procurarAcessorios(pimentaAcessoriosConf, pimentaDiferenca, pimentaData)

      procurarAcessorios(vilhenaAcessoriosConf, vilhenaDiferenca, vilhenaData)

      procurarAcessorios(cacoalAcessoriosConf, cacoalDiferenca, cacoalData)

      procurarAcessorios(setembroAcessoriosConf, setembroDiferenca, setembroData)

      procurarAcessorios(jatuaranaAcessoriosConf, jatuaranaDiferenca, jatuaranaData)

      procurarAcessorios(joseAcessoriosConf, joseDiferenca, joseData)

      procurarAcessorios(confresaAcessoriosConf, confresaDiferenca, confresaData)

      procurarAcessorios(xavantinaAcessoriosConf, xavantinaDiferenca, xavantinaData)
    }

    // Caso haja diferença conferência x sistema, exportará o arquivo
    verificarDiff(lucasDiferenca, 'Lucas', 0)

    verificarDiff(filintoDiferenca, 'Filinto', 0)

    verificarDiff(roiDiferenca, 'Rondonopolis', 0)

    verificarDiff(caceresDiferenca, 'Caceres', 0)

    verificarDiff(sorrisoDiferenca, 'Sorriso', 0)

    verificarDiff(estoqueAdmDiferenca, 'Estoque ADM', 0)

    verificarDiff(vgShoppingDiferenca, 'VG Shopping', 0)

    verificarDiff(barraDiferenca, 'Barra do Garças', 0)

    verificarDiff(coutoDiferenca, 'Couto', 1500)

    verificarDiff(primaveraDiferenca, 'Primavera', 1500)

    verificarDiff(pontesDiferenca, 'Pontes', 1500)

    verificarDiff(coliderDiferenca, 'Colider', 1500)

    verificarDiff(mirassolDiferenca, 'Mirassol', 1500)

    verificarDiff(guarantaDiferenca, 'Guarantã', 1500)

    verificarDiff(jaciaraDiferenca, 'Jaciara', 1500)

    verificarDiff(comodoroDiferenca, 'Comodoro', 1500)

    verificarDiff(altoAraguaiaDiferenca, 'Alto Araguaia', 1500)

    verificarDiff(ariquemesDiferenca, 'Ariquemes', 3000)

    verificarDiff(estoqueRoDiferenca, 'Estoque RO', 3000)

    verificarDiff(querenciaDiferenca, 'Querência', 3000)

    verificarDiff(jaruDiferenca, 'Jaru', 3000)

    verificarDiff(jiParanaDiferenca, 'Ji Paraná', 3000)

    verificarDiff(peixotoDiferenca, 'Peixoto', 3000)

    verificarDiff(rolimDiferenca, 'Rolim de Moura', 3000)

    verificarDiff(pimentaDiferenca, 'Pimenta Bueno', 3000)

    verificarDiff(vilhenaDiferenca, 'Vilhena', 4500)

    verificarDiff(cacoalDiferenca, 'Cacoal', 4500)

    verificarDiff(setembroDiferenca, '07 Setembro', 4500)

    verificarDiff(jatuaranaDiferenca, 'Jatuarana', 4500)

    verificarDiff(joseDiferenca, 'Jose Amador', 4500)

    verificarDiff(confresaDiferenca, 'Confresa', 4500)

    verificarDiff(xavantinaDiferenca, 'Xavantina', 4500)

    document.getElementById('p').setAttribute('style', 'display: block')
  }


  // Criar arquivo XLSX
  function exportToXLSX(data, nome) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, `${nome}`);
    const wbout = XLSX.write(wb, { type: 'binary', bookType: 'xlsx' });

    const s2ab = (s) => {
      const buf = new ArrayBuffer(s.length);
      const view = new Uint8Array(buf);
      for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
      return buf;
    };

    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    const fileName = `${nome}.xlsx`;
    if (typeof window !== 'undefined' && window.navigator && window.navigator.msSaveOrOpenBlob) {
      window.navigator.msSaveOrOpenBlob(blob, fileName);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
    }
  }

  const [categoria, setCategoria] = useState('Diversos')
  const categoriaChange = (e) => {
    setCategoria(e.target.value)
  }


  let [ren, setRen] = useState('')
  function setListaFaltantes() {
    setRen(() => {
      return (
        <>
          <label id="labelFileInput" htmlFor='fileInput'>2 - Escolha o arquivo .xlsx</label>
          <input style={{ display: 'none' }} type="file" id="fileInput" accept=".xlsx" onChange={estoque}></input>
          <button onClick={conferencias}>3 - Atualizar conferências</button>
          <button onClick={leitura}>4 - Verificar</button>
        </>
      )
    })
  }


  return (
    <>
      <h1>Conferência</h1>
      <p>Para iniciar o processo, selecione a categoria e clique no botão: "Verificar Lojas"</p>
      <p>Categoria:</p>
      <select name="categoria" id="categoria" value={categoria} onChange={categoriaChange}>
        <option value="Diversos">Diversos</option>
        <option value="Chips">Chips</option>
        <option value="ProdutosVivo">Produtos Vivo</option>
        <option value="Acessorios">Acessórios</option>
      </select>
      <p></p>
      <button onClick={fizeram}>1 - Verificar Lojas</button>

      {
        faltantes.length != 0
          ?
          <>
            <h4>Lojas que ainda não fizeram a conferência de {categoria}:</h4>
            <div className='divLojas'>
              {
                faltantes.map((item, index) => {
                  return (
                    <div key={index}>
                      <p className='lojas'>{item}</p>
                    </div>
                  )
                })

              }
            </div>
          </>
          :
          <></>
      }

      {ren}
      <p style={{ display: 'none' }} id='p'>Terminou! Caso não haja downloads, as conferências estão ok..</p>
    </>
  )
}

export default App