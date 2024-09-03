// LIB v0.1
/*
  Dependências
   - Swal2

  Inicializadores
   - new FormularioColunaCustom(element: string, listInternalName: string, itemId: string || null) | vai inicializar e gerar o formulário
   - new TabelaColunaCustom(element: string, listInternalName: string) | vai inicializar e gerar a tabela

  Eventos
   - formularioColunaCustomSendEvent | Evento que dispara quando o formulário criou ou atualizou o item na lista com sucesso
   - tabelaCustomEditEvent | Evento que dispara quando o botão de Editar de um item na tabela for clicado, o id do item é informado
*/

// Criar evento para o botão de enviar do formulário
const FormularioColunaCustomSendEvent = new Event('formularioColunaCustomSendEvent')

// Criar evento para o botão de editar da tabela
var TabelaCustomEditEvent

class FormularioColunaCustom {
  valores
  /*
    Função de constructor ao iniciar o formulário
    element (string) = querySelector do elemento que será inserido o formulário
    listInternalName (string) = nome interno da lista que servirá de modelo para o formulário
    itemId (string || null) = item do ID quando o formulário for um UPDATE de um item da lista, caso não for, não precisa passar valor neste parametro
  */
  constructor(element, listInternalName, itemId = null) {
    Swal.showLoading()
    this.colunas = []
    this.nomeInterno = listInternalName
    this.elemento = element
    this.itemId = null

    this.loadCampos()

    setTimeout(() => {
      if (itemId != null) {
        this.itemId = itemId
        this.preencherCampos()
      }
      Swal.close()
    }, 500);
  }

  loadCampos() {
    try {
      let url = `../_api/web/lists/getbytitle('${this.nomeInterno}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`
      let opt = {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;charset=UTF-8',
          'X-RequestDigest': document.getElementById('__REQUESTDIGEST').value
        }
      }
      fetch(url, opt).then((resp) => {
        resp.json().then((data) => {
          this.colunas = data.d.results.filter((x) => x.TypeAsString != 'Computed' && x.EntityPropertyName != 'Attachments')
          let html = `<section class="custom-main">`
          this.colunas.forEach((item) => {
            html += this.returnInputHtml(item)
          })
          html += `<div class="custom-cont-btns">
            <button class="custom-btn-back" type="button">Voltar</button>
            <button class="custom-btn-clear" type="button">Limpar</button>
            <button class="custom-btn-send" type="button">Enviar</button>
          </div></section>`
          document.querySelector(this.elemento).innerHTML = html
          document.querySelector(`${this.elemento} .custom-btn-back`).onclick = () => {
            history.back()
          }
          document.querySelector(`${this.elemento} .custom-btn-clear`).onclick = () => {
            this.limpaCampos()
          }
          document.querySelector(`${this.elemento} .custom-btn-send`).onclick = () => {
            Swal.showLoading()
            this.validate()
          }
        })
      })

    } catch (e) {
      console.error(e)
    }
  }

  preencherCampos() {
    try {
      let url = `../_api/web/lists/getbytitle('${this.nomeInterno}')/items(${this.itemId})`
      let opt = {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;charset=UTF-8',
          'X-RequestDigest': document.getElementById('__REQUESTDIGEST').value
        }
      }
      fetch(url, opt).then((resp) => {
        resp.json().then((data) => {
          this.colunas.forEach((x) => {
            switch (x.TypeAsString) {
              case 'DateTime':
                if(data.d[x.StaticName] != null) {
                  let dt = new Date(data.d[x.StaticName])
                  let str = `${dt.getFullYear()}-${((dt.getMonth() < 9) ? '0' + (dt.getMonth() + 1) : dt.getMonth())}-${dt.getDate()}`
                  document.querySelector(`.custom-main [data-col="${x.StaticName}"]`).value = str
                }
                break
              case 'Choice':
                if(data.d[x.StaticName] != null){
                  if(x.EditFormat == 0){
                    document.querySelector(`.custom-main [data-col="${x.StaticName}"]`).value = data.d[x.StaticName]
                  } else {
                    document.querySelector(`.custom-main [name="${x.StaticName}"][value="${data.d[x.StaticName]}"]`).checked = true
                  }
                }
                break
              case 'MultiChoice':
                if(data.d[x.StaticName] != null){
                  data.d[x.StaticName].results.forEach((item) => {
                    document.querySelector(`.custom-main [name="${x.StaticName}"][value="${item}"]`).checked = true
                  })
                }
                break
              default:
                document.querySelector(`.custom-main [data-col="${x.StaticName}"]`).value = data.d[x.StaticName]
            }
          })
        })
      })
    } catch (e) {
      console.log(e)
    }
  }

  returnInputHtml(col) {
    let html = `<label class="custom-cont-input">${col.Title} ${(col.Required ? '*' : '')}`
    switch (col.TypeAsString) {
      case 'Text':
        html += `<input type="text" class="custom-input" data-col="${col.StaticName}" ${(col.Required ? 'required' : '')}/>`
        break;
      case 'Choice':
        if (col.EditFormat == 0) {
          html += `<select class="custom-select" data-col="${col.StaticName}" ${(col.Required ? 'required' : '')}>
            <option value="" selected disabled>Selecione uma opção</option>`
          col.Choices.results.forEach((item) => {
            html += `<option value="${item}">${item}</option>`
          })
          html += `</select>`
        } else {
          html += `<div class="custom-cont-radio" ${(col.Required ? 'required' : '')}>`
          col.Choices.results.forEach((item) => {
            html += `<label>
              <input type="radio" name="${col.StaticName}" value="${item}" />
              ${item}
            </label>`
          })
          html += `</div>`
        }
        break
      case 'DateTime':
        html += `<input type="date" class="custom-input" data-col="${col.StaticName}" ${(col.Required ? 'required' : '')}/>`
        break;
      case 'Note':
        html += `<textarea class="custom-textarea" rows="5" data-col="${col.StaticName}" ${(col.Required ? 'required' : '')}></textarea>`
        break;
      case "MultiChoice":
        html += `<div class="custom-cont-checkbox" ${(col.Required ? 'required' : '')}>`
        col.Choices.results.forEach((item) => {
          html += `<label>
            <input type="checkbox" name="${col.StaticName}" value="${item}" />
            ${item}
          </label>`
        })
        html += `</div>`
        break
      case 'User':
        console.log(col)
        html += `<div class="custom-cont-combobox" ${(col.Required ? 'required' : '')}>
          <input type="text" data-col="${col.StaticName}" />`
        
        html += `</div>`
        break
      default:
        console.log('Tipo de campo inexistente!', col)
        // throw new Error('Tipo de campo inexistente!', col)
    }
    html += `</label>`
    return html
  }

  validate() {
    let err = 0
    let elems = document.querySelectorAll('.custom-cont-input > [required]')
    for (let el of elems) {
      if (el.classList.contains('custom-cont-radio')) {
        let name = el.children[0].children[0].attributes.name.value
        const checked = document.querySelector(`input[name="${name}"]:checked`)
        if (!checked) {
          err++
          if (!el.parentElement.children[el.parentElement.children.length - 1].classList.contains('custom-error-msg')) {
            el.parentElement.innerHTML += `<span class="custom-error-msg">Campo obrigatório<span>`
          }
        } else {
          if (el.parentElement.children[el.parentElement.children.length - 1].classList.contains('custom-error-msg')) {
            el.parentElement.children[el.parentElement.children.length - 1].remove()
          }
        }
      } else if(el.classList.contains('custom-cont-checkbox')) {
        let name = el.children[0].children[0].attributes.name.value
        const nChecked = document.querySelectorAll(`input[name="${name}"]:checked`).length
        if (nChecked == 0) {
          err++
          if (!el.parentElement.children[el.parentElement.children.length - 1].classList.contains('custom-error-msg')) {
            el.parentElement.innerHTML += `<span class="custom-error-msg">Campo obrigatório<span>`
          }
        } else {
          if (el.parentElement.children[el.parentElement.children.length - 1].classList.contains('custom-error-msg')) {
            el.parentElement.children[el.parentElement.children.length - 1].remove()
          }
        }
      } else if (el.value === null || el.value.trim() === "") {
        err++
        if (!el.parentElement.children[el.parentElement.children.length - 1].classList.contains('custom-error-msg')) {
          el.parentElement.innerHTML += `<span class="custom-error-msg">Campo obrigatório<span>`
        }
      } else {
        if (el.parentElement.children[el.parentElement.children.length - 1].classList.contains('custom-error-msg')) {
          el.parentElement.children[el.parentElement.children.length - 1].remove()
        }
      }
    }
    if (err) {
      Swal.fire({
        title: 'Erro no formulário!',
        text: 'Verifique os campos obrigatórios.',
        icon: 'error'
      })
      return
    } else {
      this.postIncluiAtualiza()
    }
  }

  getDataREST() {
    let lm = this.nomeInterno
    let obj = {
      __metadata: { type: `SP.Data.${lm.charAt(0).toUpperCase()}${lm.slice(1)}ListItem` },
    }
    this.colunas.forEach((x) => {
      if(x.TypeAsString == 'Choice' && x.EditFormat == 1){
        let v = document.querySelector(`input[name="${x.StaticName}"]:checked`)
        obj[x.StaticName] = ((v == null) ? '' : v.value)
      } else if(x.TypeAsString == 'MultiChoice'){
        let v = document.querySelectorAll(`input[name="${x.StaticName}"]:checked`)
        if(v.length == 0){
          obj[x.StaticName] = {"__metadata":{"type": "Collection(Edm.String)"}, "results": []}
        } else {
          let arr = []
          v.forEach((item) => {
            arr.push(item.value)
          })
          obj[x.StaticName] = {"__metadata":{"type": "Collection(Edm.String)"}, "results": arr}
        }
      } else if(x.TypeAsString == 'DateTime'){
        let v = document.querySelector(`.custom-main [data-col="${x.StaticName}"]`).value
        obj[x.StaticName] = ((v == '') ? null : v)
      } else {
        obj[x.StaticName] = document.querySelector(`.custom-main [data-col="${x.StaticName}"]`).value
      }

    })
    return JSON.stringify(obj)
  }

  limpaCampos() {
    let elems = document.querySelectorAll('.custom-main [data-col]')
    for (let el of elems) {
      el.value = ''
    }
  }

  postIncluiAtualiza() {
    // Faz o POST do formulário para a lista
    if (this.itemId == null) {
      try {
        let url = `../_api/web/lists/getbytitle('${this.nomeInterno}')/items`
        let obj = {
          method: `POST`,
          headers: {
            'accept': 'application/json;odata=verbose',
            "content-type": "application/json;odata=verbose",
            'X-RequestDigest': document.querySelector("#__REQUESTDIGEST").value,
          },
          body: this.getDataREST()
        }
        fetch(url, obj).then((resp) => {
          if (!resp.ok) {
            throw new Error('Erro na inserção', resp)
          }
          Swal.fire({
            title: 'Salvo!',
            text: 'Dados foram salvos com sucesso.',
            icon: 'success'
          }).then(() => {
            this.limpaCampos()
            // Ativa evento de enviar
            document.dispatchEvent(FormularioColunaCustomSendEvent)
          })
        })
      } catch (e) {
        console.error(e)
      }
    } else {
      try {
        let url = `../_api/web/lists/getbytitle('${this.nomeInterno}')/items(${this.itemId})`
        let obj = {
          method: `MERGE`,
          headers: {
            'accept': 'application/json;odata=verbose',
            "content-type": "application/json;odata=verbose",
            'X-RequestDigest': document.querySelector("#__REQUESTDIGEST").value,
            "IF-MATCH": "*",
          },
          body: this.getDataREST()
        }
        fetch(url, obj).then((resp) => {
          if (!resp.ok) {
            throw new Error('Erro na atualização', resp)
          }
          Swal.fire({
            title: 'Salvo!',
            text: 'Dados foram salvos com sucesso.',
            icon: 'success'
          }).then(() => {
            this.limpaCampos()
            this.itemId = 0
            // Ativa evento de enviar
            document.dispatchEvent(FormularioColunaCustomSendEvent)
          })
        })
      } catch (e) {
        console.log(e)
      }
    }
  }

}

/*
people picker
arquivo
*/

class TabelaColunaCustom {
  /*
    Função de constructor ao iniciar a tabela
    element (string) = querySelector do elemento que será inserido a tabela
    listInternalName (string) = nome interno da lista que servirá de modelo para a tabela
  */
  constructor(element, listInternalName) {
    this.colunas = []
    this.items = []
    this.nomeInterno = listInternalName
    this.elemento = element
    this.loadTabela()
  }

  loadTabela() {
    try {
      let urlCols = `../_api/web/lists/getbytitle('${this.nomeInterno}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`
      let opt = {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;charset=UTF-8',
          'X-RequestDigest': document.getElementById('__REQUESTDIGEST').value
        }
      }
      fetch(urlCols, opt).then((resp) => {
        resp.json().then((data) => {
          this.colunas = data.d.results.filter((x) => x.TypeAsString != 'Computed' && x.EntityPropertyName != 'Attachments')
          let urlItem = `../_api/web/lists/getbytitle('${this.nomeInterno}')/items`
          fetch(urlItem, opt).then((respItem) => {
            respItem.json().then((items) => {
              this.items = items.d.results
              let html = `<section class="custom-main"><table class="custom-tabela"><thead><tr>`
              this.colunas.forEach((x) => {
                html += `<th>${x.Title}</th>`
              })
              html += `<th></th><th></th></tr></thead><tbody>`
              this.items.forEach((x) => {
                html += `<tr>`
                this.colunas.forEach((y) => {
                  if(x[y.StaticName] == null){
                    html += `<td></td>`  
                  } else {
                    switch (y.TypeAsString) {
                      case 'MultiChoice':
                        html += `<td>${x[y.StaticName].results.join(' | ')}</td>`
                        break
                      case 'DateTime':
                        let dat = new Date(x[y.StaticName])
                        let str = `${dat.getDate()}/${dat.getMonth() + 1}/${dat.getFullYear()}`
                        html += `<td>${str}</td>`
                        break
                      default:
                        html += `<td>${((x[y.StaticName] == null) ? '' : x[y.StaticName])}</td>`
                    }
                  }
                })
                html += `<td><button type="button" class="custom-btn-edit" data-id="${x.Id}">Editar</button></td>`
                html += `<td><button type="button" class="custom-btn-delete" data-id="${x.Id}">Apagar</button></td>`
                html += `</tr>`
              })
              html += `</tbody></table></section>`
              document.querySelector(this.elemento).innerHTML = html
              document.querySelectorAll(`${this.elemento} .custom-btn-edit`).forEach((el) => {
                el.onclick = () => {
                  TabelaCustomEditEvent = new CustomEvent('tabelaCustomEditEvent', { detail: el.dataset.id })
                  document.dispatchEvent(TabelaCustomEditEvent)
                }
              })
              document.querySelectorAll(`${this.elemento} .custom-btn-delete`).forEach((el) => {
                el.onclick = () => { this.deleteItem(el.dataset.id) }
              })
            })
          })
        })
      })
    } catch (e) {
      console.log(e)
    }
  }

  deleteItem(id) {
    Swal.fire({
      title: 'Quer mesmo deletar este item?',
      text: 'Essa mudança não poderá ser revertida.',
      showCancelButton: true,
      cancelButtonText: 'Cancelar',
      confirmButtonText: 'Deletar',
      icon: 'question'
    }).then((result) => {
      if (result.isConfirmed) {
        try {
          let url = `../_api/web/lists/getbytitle('${this.nomeInterno}')/items(${id})`
          let opt = {
            method: 'DELETE',
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;charset=UTF-8',
              'X-RequestDigest': document.getElementById('__REQUESTDIGEST').value,
              'IF-MATCH': '*',
              'X-HTTP-Method': 'DELETE'
            }
          }
          fetch(url, opt).then((resp) => {
            if (!resp.ok) {
              throw new Error('Erro na exclusão!', resp)
            }
            Swal.fire('Item deletado!', '', 'success').then(() => { this.loadTabela() })
          })
        } catch (e) {
          console.log(e)
        }
      }
    })
  }
}

/*
  MUDANÇAS
  
*/
