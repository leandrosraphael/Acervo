<script>    
    $(document).ready(function(){
        for (var i = 0; i < sessionStorage.length; i++){
            var key = sessionStorage.key(i);       
            var imagem = sessionStorage.getItem(key);       
            if ((imagem != null) && (imagem.indexOf("data:image") >= 0))
            {
                $("img[data-tipo='" + key + "']").attr("src",imagem); 
                $("td[data-tipo='" + key + "']").css("display","inline");            
            }
        }
        sessionStorage.clear();
    });
</script>