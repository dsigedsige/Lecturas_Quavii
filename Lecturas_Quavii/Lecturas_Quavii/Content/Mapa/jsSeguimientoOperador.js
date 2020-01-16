
//var Pyliline;
//var Lista_marker = [];

//var Ruta = '/Lecturas'
////var Ruta = ''

//function funct_pintar_pdv(_vdata) {
//    map.removePolylines();
//    map.removeMarkers();
//    _datalinea = [];
//    v_final = 0;
//    v_medio = 0;

//    console.log('Entro Seguimiento lectura')


//    $.each(_vdata, function (key, value) {
//        v_final = key
//    });
//     $.each(_vdata, function (key, value) {
      
 
//        var _vvontentS = "";
//        var _vvontentO = "";

//        if (parseFloat(value._b) != 0 && parseFloat(value._c) != 0) {
//            _vvontentO += '<div style="width:400px;height:200px;position:relative;">';
//            _vvontentO += '<table><tr><td><strong >Suministro</strong></td><td>: ' + value._m + '</td></tr>';
//            _vvontentO += '<tr><td><strong>Medidor</strong></td><td>: ' + value._n + '</td></tr>';
//            _vvontentO += '<tr><td><strong>Dirección</strong></td><td>: ' + value._o + '</td></tr>';
//            _vvontentO += '<tr><td><strong>Operador</strong></td><td>: ' + value._f + '</td></tr>';
//            _vvontentO += '<tr><td><strong>Cliente</strong></td><td>: ' + value._p + ' </td></tr>';
//            _vvontentO += '<tr><td><strong>Lectura</strong></td><td>: ' + value._q + ' </td></tr>';
//            if (value._r == 'SI') {
//                _vvontentO += '<tr><td colspan="2" style="text-align:center;"><strong><a onclick="galeriaFotos(' + value._l + ');" href="#">Ver Foto</a></strong></td></tr>';
//            }

//            _vvontentO += '<tr><td colspan="2" style="text-align:center;"><strong><a onclick="fnModalHistorico(' + value._m + ');" href="#">Historico de Lecturas</a></strong></td></tr>';
//            _vvontentO += '</table>';
//            _vvontentO += '</div>';

//            var myLatlng = new google.maps.LatLng(value._b, value._c);

//            if (value._s != '7' && value._s != '8') {
               
//                if (key == 0) {
//                    var marcador = map.addMarker({
//                        lat: parseFloat(value._b),
//                        lng: parseFloat(value._c),
//                        title: value._f,
//                        //code: value._i,
//                        icon: Ruta + '/Content/Imagen/lectura_inicio.png',
//                       // icon: '/Content/Imagen/lectura_inicio.png',
//                        click: function (e) {
//                            map.setCenter(value._b, value._c);
//                            map.setZoom(15);
//                        },
//                        infoWindow: {
//                            content: _vvontentO
//                        }
//                    });
//                } else {
//                    var marcador = map.addMarker({
//                        lat: parseFloat(value._b),
//                        lng: parseFloat(value._c),
//                        title: value._f,
//                        //code: value._i,
//                        icon: Ruta + '/Content/Imagen/m_red.png',
//                       // icon: '/Content/Imagen/m_red.png',
//                        click: function (e) {
//                            map.setCenter(value._b, value._c);
//                            map.setZoom(15);
//                        },
//                        infoWindow: {
//                            content: _vvontentO
//                        }
//                    });
//                    //map.setZoom(15);
//                    Lista_marker.push(marcador);
//                    Lista_marker[value._a] = marcador;
//                    _datalinea.push(myLatlng);
//                }

//            } else {

//                ///key == 0 primer registro
                   
//                if (key == 0) {
                 
//                    var marcador = map.addMarker({
//                        lat: parseFloat(value._b),
//                        lng: parseFloat(value._c),
//                        title: value._f,
//                        //code: value._i,
//                        icon: Ruta + '/Content/Imagen/lectura_inicio.png',
//                        // icon: '/Content/Imagen/lectura_inicio.png',
//                        click: function (e) {
//                            map.setCenter(value._b, value._c);
//                            map.setZoom(15);
//                        },
//                        infoWindow: {
//                            content: _vvontentO
//                        }
//                    });
//                } else {

//                    var marcador = map.addMarker({
//                        lat: parseFloat(value._b),
//                        lng: parseFloat(value._c),
//                        title: value._f,
//                        //code: value._i,
//                        icon: Ruta + '/Content/Imagen/m_blue.png',
//                        //icon: '/Content/Imagen/m_blue.png',
//                        click: function (e) {
//                            map.setCenter(value._b, value._c);
//                            map.setZoom(15);
//                        },
//                        infoWindow: {
//                            content: _vvontentO
//                        }
//                    });
//                    //map.setZoom(15);
//                    Lista_marker.push(marcador);
//                    Lista_marker[value._a] = marcador;
//                    _datalinea.push(myLatlng);
//                }
//            }
            

//            if (key == v_final) {
//                var marcador = map.addMarker({
//                    lat: parseFloat(value._b),
//                    lng: parseFloat(value._c),
//                    title: value._f,
//                    //code: value._i,
//                    icon: Ruta + '/Content/Imagen/lectura_fin.png',
//                    //icon: '/Content/Imagen/lectura_fin.png',
//                    click: function (e) {
//                        map.setCenter(value._b, value._c);
//                        map.setZoom(15);
//                    },
//                    infoWindow: {
//                        content: _vvontentO
//                    }
//                });
//                //map.setZoom(15);
//                Lista_marker.push(marcador);
//                Lista_marker[value._a] = marcador;
//                _datalinea.push(myLatlng);
//            }


//        }  // fin msg
//    });  // fin bucle 

//    Pyliline = map.drawPolyline({
//        path: _datalinea,
//        strokeColor: '#337AB7',
//        strokeOpacity: 0.6,
//        strokeWeight: 6
//    });

//}

/// seguimiento Operario




//function funct_pintar_pdv_seguimiento_operario2(_vdata) {
//    console.log('entro a la funcion')

//    _datalinea = [];
//    v_final = 0;
//    v_medio = 0;

//    console.log('entro a la funcion')
//    console.log(Ruta + '/Content/Imagen/operario.png')
//    console.log(_vdata)


//    $.each(_vdata, function (key, value) {
//        v_final = key
//    });
//    v_medio = v_final / 2;
//    v_medio = Math.floor(v_medio);

//    $.each(_vdata, function (key, value) {

//    var _vvontentO = "";
   
//            if (parseFloat(value._c) != 0 && parseFloat(value._d) != 0) {
//                var myLatlng = new google.maps.LatLng(value._c, value._d);


//                if (key == 0) {
             
//                    console.log('Primer Operario')

//                    var marcador = map.addMarker({
//                        lat: parseFloat(value._c),
//                        lng: parseFloat(value._d),
//                        title: value._e,
//                        //code: value._i,
//                        icon: Ruta + '/Content/Imagen/operario.png',
//                        //icon: '/Content/Imagen/operario.png',
//                        click: function (e) {
//                            map.setCenter(value._c, value._d);
//                            map.setZoom(15);
//                        }
//                    });
//                    Lista_marker.push(marcador);
//                    Lista_marker[value._a] = marcador;
//                    _datalinea.push(myLatlng);

//                    console.log('entro el primer operario seguimiento')

//                }
//                else {

//                    if (key == v_final) {
                                   
//                        var marcador = map.addMarker({
//                            lat: parseFloat(value._c),
//                            lng: parseFloat(value._d),
//                            //title: value._a,
//                            //code: value._i,
//                            icon: Ruta + '/Content/Imagen/operario_final.png',
//                            //icon: '/Content/Imagen/operario_final.png',
//                            click: function (e) {
//                                map.setCenter(value._d, value._d);
//                                map.setZoom(15);
//                            }
//                        });
//                        Lista_marker.push(marcador);
//                        Lista_marker[value._a] = marcador;
//                        _datalinea.push(myLatlng);
//                    } else {
 

//                        if (value._f >= 5) {
                            
//                            /// Validando si posee paradas
                        
//                            if (parseFloat(value._c) != 0 && parseFloat(value._d) != 0) {
//                                _vvontentO += '<div style="width:400px;height:100px;position:relative;">';
//                                _vvontentO += '<table><tr><td><strong >Operario</strong></td><td>: ' + value._b + '</td></tr>';
//                                _vvontentO += '<tr><td><strong>Tiempo Parada</strong></td><td>: ' + value._f + '</td></tr>';
//                                _vvontentO += '<tr><td><strong>Hora de inicio parada</strong></td><td>: ' + value._g + '</td></tr>';
//                                _vvontentO += '<tr><td><strong>Hora fin de parada</strong></td><td>: ' + value._h + '</td></tr>';
//                                _vvontentO += '</table>';
//                                _vvontentO += '</div>';
//                            } 
//                            var marcador = map.addMarker({
//                                lat: parseFloat(value._c),
//                                lng: parseFloat(value._d),
//                                title: value._b,
//                                //code: value._i,
//                                icon: Ruta + '/Content/Imagen/ope_paro.png',
//                                //icon: '/Content/Imagen/ope_paro.png',
//                                click: function (e) {
//                                    map.setCenter(value._c, value._d);
//                                    map.setZoom(15);
//                                },
//                                infoWindow: {
//                                    content: _vvontentO
//                                }

//                            });
//                            //map.setZoom(15);
//                            Lista_marker.push(marcador);
//                            Lista_marker[value._a] = marcador;
//                            _datalinea.push(myLatlng);

//                        }
//                        else {

//                            var marcador = map.addMarker({
//                                lat: parseFloat(value._c),
//                                lng: parseFloat(value._d),
//                                title: value._b,
//                                //code: value._i,
//                                icon: Ruta + '/Content/Imagen/seg_operario.png',
//                                //icon: '/Content/Imagen/seg_operario.png',
//                                click: function (e) {
//                                    map.setCenter(value._c, value._d);
//                                    map.setZoom(15);
//                                },
//                            });
//                            //map.setZoom(15);
//                            Lista_marker.push(marcador);
//                            Lista_marker[value._a] = marcador;
//                            _datalinea.push(myLatlng);

//                        }                  

//                    }                
//                }
//            }
//          /// fin de msg
//    });  // fin de bucle

//        Pyliline = map.drawPolyline({
//            path: _datalinea,
//            strokeColor: '#19B200',
//            strokeOpacity: 0.6,
//            strokeWeight: 6
//        });

//    }  // fin de la funcion 
   




































// anterior

//function funct_pintar_pdv_seguimiento_operario(_vdata) {

//    //console.log(_vdata)
//    //return;

//    _datalinea = [];
//    v_final = 0;
//    v_medio = 0;

//    $.each(_vdata, function (key, value) {
//        v_final = key
//    });
//    v_medio = v_final / 2;
//    v_medio = Math.floor(v_medio);

//    $.each(_vdata, function (key, value) {
        
        
//        if (parseFloat(value._b) != 0 && parseFloat(value._c) != 0) {
//            var myLatlng = new google.maps.LatLng(value._b, value._c);


//            if (key == 0) {
//                var marcador = map.addMarker({


//                    lat: parseFloat(value._b),
//                    lng: parseFloat(value._c),
//                    //title: value._a,
//                    //code: value._i,
//                    //icon: '/Lecturas/Content/Imagen/operario.png',
//                    icon: '/Content/Imagen/operario.png',
//                    click: function (e) {
//                        map.setCenter(value._b, value._c);
//                        map.setZoom(15);
//                    }
//                });
//                Lista_marker.push(marcador);
//                Lista_marker[value._a] = marcador;
//                _datalinea.push(myLatlng);

//            } else {
//                if (key != v_final && key != v_medio && key != 0 && value._aa == "1") {
//                //if (value._aa == "1") {
//                    var marcador = map.addMarker({
//                        lat: parseFloat(value._b),
//                        lng: parseFloat(value._c),
//                        icon: '/Content/Imagen/ope_paro.png',
//                       // icon: '/Lecturas/Content/Imagen/ope_paro.png',
//                        click: function (e) {
//                            map.setCenter(value._b, value._c);
//                            map.setZoom(15);
//                        }
//                    });
//                    Lista_marker.push(marcador);
//                    Lista_marker[value._a] = marcador;
//                    _datalinea.push(myLatlng);
//                } else {
//                    _datalinea.push(myLatlng);
//                }

//            }

//            if (key == v_medio) {
//                var marcador = map.addMarker({
//                    lat: parseFloat(value._b),
//                    lng: parseFloat(value._c),
//                    //title: value._a,
//                    //code: value._i,
//                    //icon: '/Lecturas/Content/Imagen/seg_operario.png',
//                    icon: '/Content/Imagen/seg_operario.png',
//                    click: function (e) {
//                        map.setCenter(value._b, value._c);
//                        map.setZoom(15);
//                    }
//                });
//                Lista_marker.push(marcador);
//                Lista_marker[value._a] = marcador;
//                _datalinea.push(myLatlng);
//            }

//            if (key == v_final) {
//                var marcador = map.addMarker({
//                    lat: parseFloat(value._b),
//                    lng: parseFloat(value._c),
//                    //title: value._a,
//                    //code: value._i,
//                   // icon: '/Lecturas/Content/Imagen/operario_final.png',
//                    icon: '/Content/Imagen/operario_final.png',
//                    click: function (e) {
//                        map.setCenter(value._b, value._c);
//                        map.setZoom(15);
//                    }
//                });
//                Lista_marker.push(marcador);
//                Lista_marker[value._a] = marcador;
//                _datalinea.push(myLatlng);
//            }


//        }
//    });

//    Pyliline = map.drawPolyline({
//        path: _datalinea,
//        strokeColor: '#19B200',
//        strokeOpacity: 0.6,
//        strokeWeight: 6
//    });

//}





function fnModalHistorico(cod) {
    $('#modal-historico').modal('show');
    $.ajax({
        async: true,
        beforeSend: function (xhr) {
        },
        url: Ruta + '/Asigna_Lectura/JsonListaHistorico',
        //url: '/Asigna_Lectura/JsonListaHistorico',
        type: 'POST',
        dataType: 'json',
        contentType: 'application/json; charset=utf-8',
        data: JSON.stringify({
            //__a: $contenedor.join(',')
            __a: cod
        }),
        success: function (__a) {
            $('#table-historico').bootstrapTable('load', __a);
            $('#table-historico').bootstrapTable('resetView');
        },
        error: function (xhr) {
            alert("Algo salió mal, por favor intente de nuevo.");
        }
    });


};


function galeriaFotos(idLectura) {
    $('#fotos').modal();
    $.ajax({
        async: true,
        beforeSend: function (xhr) { },
        url: Ruta + '/Envia_Trabajos_Cli/FotosLectura',
        //url: '/Envia_Trabajos_Cli/FotosLectura',
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
            $.each(response, function (i, v) {
                console.log(i);
                if (i === 0) {
                    $DivCorre.append('<div id="x" class="item active"><img id="img-rota" src="' + $ruta + '/' + response[i]._r + '" alt="..." style="width: 320px; height: auto; text-align: center;" /></div>');
                } else {
                    $DivCorre.append('<div id="x" class="item"><img id="img-rota" src="' + $ruta + '/' + response[i]._r + '" alt="..." style="width: 320px; height: auto; text-align: center;" /></div>');
                }
            });
            //$('.img-foto').attr('src', $ruta + '/' + __s._h);
        },
        error: function (xhr) {
            alert('Algo salió mal, por favor intente de nuevo.');
        }
    });
}