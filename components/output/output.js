const block_output = document.querySelector('details.output');
block_output.addEventListener('click', (e) => {
 if(e.target.id == 'btn_clear_output'){
  document.querySelector('.output__field').innerHTML = '';
 }
})