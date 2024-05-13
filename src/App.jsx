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
    if (categoria == 'Outros') {
      relatorios.forEach(e => {
        if (e.categoria == 'Outros') {
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

  function pshConf(dataConf, acessoriosConf, e) {
    if (e.categoria == 'Outros') {
      dataConf.push(e.text)
    } else if (e.categoria == 'Acessorios') {
      acessoriosConf.push(e.text)
    }
  }

  function conferencias() {
    relatorios.forEach(e => {
      if (e.loja == 'MS LUCAS') {
        pshConf(lucasDataConf, lucasAcessoriosConf, e)
      }

      if (e.loja == 'MS FILINTO') {
        pshConf(filintoDataConf, filintoAcessoriosConf, e)
      }

      if (e.loja == 'MS ROI LLA') {
        pshConf(roiDataConf, roiAcessoriosConf, e)
      }

      if (e.loja == 'MS CÁCERES') {
        pshConf(caceresDataConf, caceresAcessoriosConf, e)
      }

      if (e.loja == 'MS SORRISO') {
        pshConf(sorrisoDataConf, sorrisoAcessoriosConf, e)
      }

      if (e.loja == 'ESTOQUE ADM') {
        pshConf(estoqueAdmDataConf, estoqueAdmAcessoriosConf, e)
      }

      if (e.loja == 'MS VG Shopping') {
        pshConf(vgShoppingDataConf, vgShoppingAcessoriosConf, e)
      }

      if (e.loja == 'MS BARRA DO GARÇAS') {
        pshConf(barraDataConf, barraAcessoriosConf, e)
      }

      if (e.loja == 'MS COUTO') {
        pshConf(coutoDataConf, coutoAcessoriosConf, e)
      }

      if (e.loja == 'MS PRIMAVERA DO LESTE') {
        pshConf(primaveraDataConf, primaveraAcessoriosConf, e)
      }

      if (e.loja == 'MS PONTES E LACERDA') {
        pshConf(pontesDataConf, pontesAcessoriosConf, e)
      }

      if (e.loja == 'MS COLIDER') {
        pshConf(coliderDataConf, coliderAcessoriosConf, e)
      }

      if (e.loja == 'MS MIRASSOL') {
        pshConf(mirassolDataConf, mirassolAcessoriosConf, e)
      }

      if (e.loja == 'MS GUARANTÃ DO NORTE') {
        pshConf(guarantaDataConf, guarantaAcessoriosConf, e)
      }

      if (e.loja == 'MS JACIARA') {
        pshConf(jaciaraDataConf, jaciaraAcessoriosConf, e)
      }

      if (e.loja == 'MS COMODORO') {
        pshConf(comodoroDataConf, comodoroAcessoriosConf, e)
      }

      if (e.loja == 'MS ALTO ARAGUAIA') {
        pshConf(altoAraguaiaDataConf, altoAraguaiaAcessoriosConf, e)
      }

      if (e.loja == 'MS ARIQUEMES') {
        pshConf(ariquemesDataConf, ariquemesAcessoriosConf, e)
      }

      if (e.loja == 'ESTOQUE ADM RO') {
        pshConf(estoqueRoDataConf, estoqueRoAcessoriosConf, e)
      }

      if (e.loja == 'MS QUERÊNCIA') {
        pshConf(querenciaDataConf, querenciaAcessoriosConf, e)
      }

      if (e.loja == 'MS JARU') {
        pshConf(jaruDataConf, jaruAcessoriosConf, e)
      }

      if (e.loja == 'MS JI-PARANA') {
        pshConf(jiParanaDataConf, jiParanaAcessoriosConf, e)
      }

      if (e.loja == 'MS PEIXOTO DE AZEVEDO') {
        pshConf(peixotoDataConf, peixotoAcessoriosConf, e)
      }

      if (e.loja == 'MS ROLIM DE MOURA') {
        pshConf(rolimDataConf, rolimAcessoriosConf, e)
      }

      if (e.loja == 'MS PIMENTA BUENO') {
        pshConf(pimentaDataConf, pimentaAcessoriosConf, e)
      }

      if (e.loja == 'MS VILHENA') {
        pshConf(vilhenaDataConf, vilhenaAcessoriosConf, e)
      }

      if (e.loja == 'MS CACOAL') {
        pshConf(cacoalDataConf, cacoalAcessoriosConf, e)
      }

      if (e.loja == 'MS PV - 07 SETEMBRO') {
        pshConf(setembroDataConf, setembroAcessoriosConf, e)
      }

      if (e.loja == 'MS PV - JATUARANA') {
        pshConf(jatuaranaDataConf, jatuaranaAcessoriosConf, e)
      }

      if (e.loja == 'MS PV - JOSE AMADOR') {
        pshConf(joseDataConf, joseAcessoriosConf, e)
      }

      if (e.loja == 'MS CONFRESA') {
        pshConf(confresaDataConf, confresaAcessoriosConf, e)
      }

      if (e.loja == 'MS NOVA XAVANTINA') {
        pshConf(xavantinaDataConf, xavantinaAcessoriosConf, e)
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
    // Quando OUTROS estiver selecionado
    if (categoria == 'Outros') {
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

      // Quando ACESSORIOS tiver selecionado
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

  const [categoria, setCategoria] = useState('Acessorios')
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
        <option value="Acessorios">Acessórios</option>
        <option value="Outros">Outros</option>
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