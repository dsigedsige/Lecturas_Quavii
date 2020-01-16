 
var Pyliline;
var Lista_marker = [];
var fechaMovil;

function funct_pintar_pdv(_vdata, fecha_movil, servicio) {


    fechaMovil = fecha_movil;
    map.removePolylines();
    map.removeMarkers();
    _datalinea = [];


    $.each(_vdata, function (key, value) {
        //if (value._l == 1) {
        var _vvontentS = "";
        var _vvontentO = "";
        if (parseFloat(value._b) != 0 && parseFloat(value._c) != 0) {
            _vvontentO += '<div style="height: 100px;width:300px;height:110px;position:relative;">';
            _vvontentO += '<center><h4><b>Ubicacion de Operarios</b></h4></center>';
            _vvontentO += '<table><tr><td><strong >Operador</strong></td><td>: ' + value._f + '</td></tr>';
            _vvontentO += '<tr><td><strong>Total</strong></td><td>: ' + value._h + '</td></tr>';
            _vvontentO += '<tr><td><strong>Realizado</strong></td><td>: ' + value._i + '</td></tr>';
            _vvontentO += '<tr><td><strong>Pendiente</strong></td><td>: ' + value._j + '</td></tr>';
            _vvontentO += '<tr><td><strong>% Avance</strong></td><td>: ' + value._k + ' %</td></tr>';        
            if (servicio == '6' || servicio ==6) {
                _vvontentO += '<tr><td colspan="2" style="text-align:center;"><strong><a onclick="galeriaFotos(' + value._a + ');" href="#">Ver Selfie</a></strong></td></tr>';
            }
            _vvontentO += '</table>';
            _vvontentO += '</div>';


            //_vvontentS += '<div style="height: 100px;width:300px;height:130px;position:relative;">';
            //_vvontentS += '<table><tr><td><strong >Operador</strong></td><td>: ' + value._f + '</td></tr>';
            //_vvontentS += '<tr><td><strong>Dirección</strong></td><td>:</td></tr>';
            //_vvontentS += '<tr><td colspan="2" style="text-align:center;"><strong><a onclick="fnListaLecturasPendientes()" href="#">Ver Resumen de Operadores</a></strong></td></tr>';
            //_vvontentS += '</table>';
            //_vvontentS += '</div>';

            var myLatlng = new google.maps.LatLng(value._b, value._c);

            if (value._g == 'S') {
                if (value._w == '0') {
                    var marcador = map.addMarker({
                        lat: parseFloat(value._b),
                        lng: parseFloat(value._c),
                        title: value._f,
                        icon: '../Content/Imagen/supervisor_off.png',
                        //icon: '/Content/Imagen/supervisor_off.png',

                        click: function (e) {
                            map.setCenter(value._b, value._c);
                           // map.setZoom(15);
                        },
                        infoWindow: {
                            content: _vvontentS
                        }
                    });
                    //map.setZoom(15);
                    Lista_marker.push(marcador);
                    Lista_marker[value._a] = marcador;
                    //_datalinea.push(myLatlng);
                } else {
                    var marcador = map.addMarker({
                        lat: parseFloat(value._b),
                        lng: parseFloat(value._c),
                        title: value._f,
                        //code: value._i,
                        icon: '../Content/Imagen/supervisor_on.png',
                        //icon: '/Content/Imagen/supervisor_on.png',

                        click: function (e) {
                            map.setCenter(value._b, value._c);
                            //map.setZoom(15);
                        },
                        infoWindow: {
                            content: _vvontentS
                        }
                    });
                    //map.setZoom(15);
                    Lista_marker.push(marcador);
                    Lista_marker[value._a] = marcador;
                    //_datalinea.push(myLatlng);
                }
            } else if (value._g == 'O') {
                if (value._w == '0') {
                    var marcador = map.addMarker({
                        lat: parseFloat(value._b),
                        lng: parseFloat(value._c),
                        title: value._f,
                        //code: value._i,
                    //*Cambiar
                        icon: '../Content/Imagen/operario_off.png',
                       //icon: '/Content/Imagen/operario_off.png',

                        click: function (e) {
                            map.setCenter(value._b, value._c);
                            //map.setZoom(15);
                        },
                        infoWindow: {
                            content: _vvontentO
                        }
                    });
                    //map.setZoom(15);
                    Lista_marker.push(marcador);
                    Lista_marker[value._a] = marcador;
                    //_datalinea.push(myLatlng);
                } else {
                    var marcador = map.addMarker({
                        lat: parseFloat(value._b),
                        lng: parseFloat(value._c),
                        title: value._f ,
                        icon: '../Content/Imagen/operario_on.png',
                        //icon: '/Content/Imagen/operario_on.png',

                        click: function (e) {
                            map.setCenter(value._b, value._c);
                            //map.setZoom(15);
                        },
                        infoWindow: {
                            content: _vvontentO
                        }
                    });
                    //map.setZoom(15);
                    Lista_marker.push(marcador);
                    Lista_marker[value._a] = marcador;
                    //_datalinea.push(myLatlng);
                }
            } else if (value._g == 'J') { //JEFE
                if (value._w == '0') {
                    var marcador = map.addMarker({
                        lat: parseFloat(value._b),
                        lng: parseFloat(value._c),
                        title: value._f,
                        icon: '../Content/Imagen/jefe_off.png',
                        //icon: '/Content/Imagen/jefe_off.png',
                        click: function (e) {
                            map.setCenter(value._b, value._c);
                            //map.setZoom(15);
                        },
                        infoWindow: {
                            content: _vvontentO
                        }
                    });
                    //map.setZoom(15);
                    Lista_marker.push(marcador);
                    Lista_marker[value._a] = marcador;
                    //_datalinea.push(myLatlng);
                } else {
                    var marcador = map.addMarker({
                        lat: parseFloat(value._b),
                        lng: parseFloat(value._c),
                        title: value._f ,  
                        icon: '../Content/Imagen/jefe_on.png',
                        //icon: '/Content/Imagen/jefe_on.png',

                        click: function (e) {
                            map.setCenter(value._b, value._c);
                           // map.setZoom(15);
                        },
                        infoWindow: {
                            content: _vvontentO
                        }
                    });
                    //map.setZoom(15);
                    Lista_marker.push(marcador);
                    Lista_marker[value._a] = marcador;
                    //_datalinea.push(myLatlng);
                }
            } else { //CLIENTE - EDELNOR
                if (value._w == '0') {
                    var marcador = map.addMarker({
                        lat: parseFloat(value._b),
                        lng: parseFloat(value._c),
                        title: value._f,
                        icon: '../Content/Imagen/cliente_off.png',
                        //icon: '/Content/Imagen/cliente_off.png',

                        click: function (e) {
                            map.setCenter(value._b, value._c);
                            //map.setZoom(15);
                        },
                        infoWindow: {
                            content: _vvontentO
                        }
                    });
                    //map.setZoom(15);
                    Lista_marker.push(marcador);
                    Lista_marker[value._a] = marcador;
                    //_datalinea.push(myLatlng);
                } else {
                    var marcador = map.addMarker({
                        lat: parseFloat(value._b),
                        lng: parseFloat(value._c),
                        title: value._f,
                        icon: '../Content/Imagen/cliente_on.png',
                        //icon: '/Content/Imagen/cliente_on.png',

                        click: function (e) {
                            map.setCenter(value._b, value._c);
                            //map.setZoom(15);
                        },
                        infoWindow: {
                            content: _vvontentO
                        }
                    });
                    //map.setZoom(15);
                    Lista_marker.push(marcador);
                    Lista_marker[value._a] = marcador;
                    //_datalinea.push(myLatlng);

                }
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


function fnListaLecturasPendientes() {
    $('#modal-pendientes').modal('show');
    $.ajax({
        async: true,
        beforeSend: function (xhr) { $('.sige-load').show(); },
        url: 'JsonUbicacion_OperariosGPS',
        type: 'POST',
        dataType: 'json',
        contentType: 'application/json; charset=utf-8',
        data: JSON.stringify({
            __a: $("#_supervisor").val(),
            __b: $("#_fechaAsigna").val()
        }),
        success: function (__a) {
            $('#table-pendientes').bootstrapTable('load', __a);
            $('#table-pendientes').bootstrapTable('resetView');

        },
        complete: function () {
            $('.sige-load').hide();
        },
        error: function (xhr) {
            $('.sige-load').hide();
            alert("Algo salió mal, por favor intente de nuevo.");
        }
    });


};

function fnDescargaOperariosSupervisor() {
    $ajaxJson(
        'DescargaOperariosSupervisor',
        function () {
            return {
                __a: JSON.stringify($('#table-pendientes').bootstrapTable('getData'))
            };
        },
        function (__s) {
            window.open('/Calidda/Temp/' + __s.__a, '_blank');
            //window.open('/Temp/' + __s.__a, '_blank');
        }
    );
}


var degrees = 0;
function rotateImg() {
    degrees += 90;
    console.log($('#imgRotate'));
    $("#imgRotate").rotate(degrees);
}

 

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


function galeriaFotos(id_operario) {
    $.ajax({
        async: true,
        beforeSend: function (xhr) { },
        url: '../Envia_Trabajos_Cli/FotoSelfie_Reparto',
        type: 'POST',
        dataType: 'json',
        data: {
            fecha: fechaMovil,
            id_operario: id_operario
        },
        success: function (response) {
            if (response.length > 0) {
                $("div").remove("#x");
                $("div").remove("#y");
                var $DivCorre = $('#corre');
                var $DivOl = $('#ol');
                listPhotos = [];

                $.each(response, function (i, v) {
                    listPhotos.push({
                        id: i,
                        url: response[i].RutaFoto
                    })
                });

                $('#fotos').modal();
                $DivCorre.append('<div id="x" class="item active"><img id="foto1" src="' + listPhotos[0].url + '" alt="..." style="width: 550px; height: auto; text-align: center;" /></div>');
            } else {
                new PNotify({
                    title: 'Sistemas',
                    text: 'El Operario aun no tiene foto de Inicio de Trabajo..',
                    type: 'error'
                });
                return;
            }
        },
        error: function (xhr) {
            alert('Algo salió mal, por favor intente de nuevo.');
        }
    });
}

