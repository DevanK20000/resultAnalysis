from django.conf import settings
from django.http import HttpResponse
from django.shortcuts import render, redirect
from django.contrib import messages
from django.contrib.auth.models import User
from django.contrib.auth import authenticate, login, logout
import pandas as pd
from base import extract
from base.analysis_mcasem1 import genrateAnalysis_MCAsem1
from base.analysis_mcasem2 import genrateAnalysis_MCAsem2
from base.analysis_mcasem3 import genrateAnalysis_MCAsem3
from base.class_analysis import class_analysis
from .models import FileUpload
# Create your views here.


def loginPage(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')

        try:
            user = User.objects.get(username=username)
        except:
            messages.error(request, 'User does not exists')

        user = authenticate(request, username=username, password=password)
        if user is not None:
            login(request, user)
            return redirect('home')
        else:
            messages.error(request, 'Invalid credentials')
    context = {}
    if request.user.is_authenticated:
        return redirect('home')
    else:
        return render(request, 'base/login.html', context)


def logoutUser(request):
    logout(request)
    return redirect('login')


def home(request):

    if request.method == "POST":
        messages.get_messages(request).used = True
        pdf = request.FILES.get('pdf')
        excel = request.FILES.get('excel')
        if pdf != None:
            file = request.FILES["pdf"]
            document = FileUpload.objects.create(file=file)
            document.save()
            file2 = extract.extract(pdf=document.file.path,
                                    type=request.POST.get('select'))
            file2 = file2.to_json()
            request.session['file'] = file2
            request.session['class'] = request.POST.get('select')
            return redirect('table')

        elif excel != None:
            file = request.FILES["excel"]
            document = FileUpload.objects.create(file=file)
            document.save()
            file2 = class_analysis(
                excel=document.file.path, class_name=request.POST.get('analyze_select'))

            response = HttpResponse(open(file2,
                                         'rb').read(), status=200, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            response['Content-Disposition'] = f'attachment; filename="{request.POST.get("analyze_select")+".xlsx"}"'
            return response

        else:
            messages.error(request, 'Please select a file')

    if request.user.is_authenticated:
        return render(request, 'base/home.html')
    else:
        return redirect('login')


def table(request):
    if request.user.is_authenticated:
        file = request.session['file']
        file = pd.read_json(file)
        file = pd.DataFrame(file)
        file_html = file.to_html(table_id='dataTable',
                                 classes='table table-bordered')

        return render(request, 'base/table.html', {'file_html': file_html})
    else:
        return redirect('login')


def download(request, pk):
    if request.user.is_authenticated:
        file = request.session['file']
        file = pd.read_json(file)
        if pk == 'CSV':
            file = pd.DataFrame(file)
            response = HttpResponse(file.to_csv(
                index=False), status=200, content_type="text/csv")
            response['Content-Disposition'] = f'attachment; filename="{request.session["class"]+".csv"}"'
            return response
        elif pk == 'EXCEL':
            pd.DataFrame(file).to_excel(
                settings.MEDIA_ROOT + "\\"+request.session['class']+'.xlsx')
            if request.session['class'] == 'mca_sem1':
                genrateAnalysis_MCAsem1(
                    settings.MEDIA_ROOT + "\\"+request.session['class']+'.xlsx', "Sheet1", "SEM1 ANALYSIS Reg. Both shifts")
            if request.session['class'] == 'mca_sem2':
                genrateAnalysis_MCAsem2(
                    settings.MEDIA_ROOT + "\\"+request.session['class']+'.xlsx',  "Sheet1", "SEM2 ANALYSIS Reg. Both shifts")
            if request.session['class'] == 'mca_sem3':
                genrateAnalysis_MCAsem3(
                    settings.MEDIA_ROOT + "\\"+request.session['class']+'.xlsx', "Sheet1", "SEM3 ANALYSIS Reg. Both shifts")

            response = HttpResponse(open(settings.MEDIA_ROOT + "\\"+request.session['class']+'.xlsx',
                                         'rb').read(), status=200, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            response['Content-Disposition'] = f'attachment; filename="{request.session["class"]+".xlsx"}"'
            return response
    else:
        return redirect('login')


def page_not_found_view(request, exception):
    return render(request, '404.html', status=404)
