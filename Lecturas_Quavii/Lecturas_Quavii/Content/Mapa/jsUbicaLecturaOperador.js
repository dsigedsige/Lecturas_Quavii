
var Pyliline;
var Lista_marker = [];

function funct_pintar_pdv(_vdata) {
    map.removePolylines();
    map.removeMarkers();
    _datalinea = [];

    $.each(_vdata, function (key, value) {

        //if (value._l == 1) {
        var _vvontentS = "";
        var _vvontentO = "";
        if (parseFloat(value._b) != 0 && parseFloat(value._c) != 0) {
            _vvontentO += '<div style="width:480px;height:200px;position:relative;">';
            _vvontentO += '<table><tr><td id="idSuministro"><strong >Suministro</strong></td><td>: ' + value._m + '</td></tr>';
            _vvontentO += '<tr><td><strong>Medidor</strong></td><td>: ' + value._n + '</td></tr>';
            _vvontentO += '<tr><td><strong>Dirección</strong></td><td>: ' + value._o + '</td></tr>';
            _vvontentO += '<tr><td><strong>Operador</strong></td><td>: ' + value._f + '</td></tr>';
            _vvontentO += '<tr><td><strong>Cliente</strong></td><td>: ' + value._p + ' </td></tr>';
            _vvontentO += '<tr><td><strong>Lectura</strong></td><td>: ' + value._q + ' </td></tr>';
            _vvontentO += '<tr><td><strong>Fecha envío movil</strong></td><td>: ' + value._u + ' </td></tr>';
            _vvontentO += '<tr><td><strong>Fecha de lectura</strong></td><td>: ' + value._t + ' </td></tr>';
            if (value._r == 'SI') {
                _vvontentO += '<tr><td colspan="2" style="text-align:center;"><strong><a onclick="galeriaFotos(' + value._l + ');" href="#">Ver Foto</a></strong></td></tr>';
            }

            
            //_vvontentO += '<tr><td colspan="2" style="text-align:center;"><strong><a onclick="fnModalHistorico(' + value._m + ');" href="#">Historico de Lecturas</a></strong></td></tr>';
            _vvontentO += '</table>';
            _vvontentO += '</div>';

            var myLatlng = new google.maps.LatLng(value._b, value._c);

            if (value._s != '7' && value._s != '8' ) {
                var marcador = map.addMarker({
                    lat: parseFloat(value._b),
                    lng: parseFloat(value._c),
                    title: value._f,
                    icon: '../Content1/Imagen/m_red.png',
                    //icon: '/Content/Imagen/m_red.png',
                    click: function (e) {
                        map.setCenter(value._b, value._c);
                    },
                    infoWindow: {
                        content: _vvontentO
                    }
                });
                Lista_marker.push(marcador);
                Lista_marker[value._a] = marcador;
            } else {
                var marcador = map.addMarker({
                    lat: parseFloat(value._b),
                    lng: parseFloat(value._c),
                    title: value._f,
                    icon: '../Content/Imagen/m_blue.png',
                    //icon: '/Content/Imagen/m_blue.png',
                    click: function (e) {
                        map.setCenter(value._b, value._c);
                       // map.setZoom(15);
                    },
                    infoWindow: {
                        content: _vvontentO
                    }
                });
                Lista_marker.push(marcador);
                Lista_marker[value._a] = marcador;
                //_datalinea.push(myLatlng);
            }

        }
    });

    Pyliline = map.drawPolyline({
        path: _datalinea,
        strokeColor: '#131540',
        strokeOpacity: 0.6,
        strokeWeight: 6
    });

}

//function galeriaFotos(idLectura) {
//    $('#fotos').modal();
//    $.ajax({
//        async: true,
//        beforeSend: function (xhr) { },
//        url: '../Envia_Trabajos_Cli/FotosLectura',
//        type: 'POST',
//        dataType: 'json',
//        data: {
//            __a: idLectura
//        },
//        success: function (response) {
//            $("div").remove("#x");
//            $("div").remove("#y");
//            var $DivCorre = $('#corre');
//            var $DivOl = $('#ol');
//            $.each(response, function (i, v) {
//                console.log(i);
//                if (i === 0) {
//                    $DivCorre.append('<div id="x" class="item active"><img id="img-rota" src="' + $ruta + '/' + response[i]._r + '" alt="..." style="width: 320px; height: auto; text-align: center;" /></div>');
//                } else {
//                    $DivCorre.append('<div id="x" class="item"><img id="img-rota" src="' + $ruta + '/' + response[i]._r + '" alt="..." style="width: 320px; height: auto; text-align: center;" /></div>');
//                }
//            });
//            //$('.img-foto').attr('src', $ruta + '/' + __s._h);
//        },
//        error: function (xhr) {
//            alert('Algo salió mal, por favor intente de nuevo.');
//        }
//    });
//}


var listPhotos = [];
function previusPhoto() {
    var photo1 = document.getElementById('foto1');
    photo1.src = listPhotos[0].url;
}
function nextPhoto() {
    var photo1 = document.getElementById('foto1');
    if (listPhotos.length > 1) {
        photo1.src = listPhotos[1].url;
    }
}


function galeriaFotos(idLectura) {
    $('#fotos').modal();

    $.ajax({
        async: true,
        beforeSend: function (xhr) { },
        url: '../Envia_Trabajos_Cli/FotosLectura',
        type: 'POST',
        dataType: 'json',
        data: {
            __a: idLectura
        },
        success: function (response) {

            $("div").remove("#x");
            $("div").remove("#y");
            var $DivCorre = $('#corre');
            var $DivOl = $('#ol');
            listPhotos = [];
            $.each(response, function (i, v) {
                listPhotos.push({
                    id: i,
                    url: $ruta + '/' + response[i]._r
                })
            });
            $DivCorre.append('<div id="x" class="item active"><img id="foto1" src="' + listPhotos[0].url + '" alt="..." style="width: 320px; height: auto; text-align: center;" /></div>');

        },
        error: function (xhr) {
            alert('Algo salió mal, por favor intente de nuevo.');
        }
    });
}
