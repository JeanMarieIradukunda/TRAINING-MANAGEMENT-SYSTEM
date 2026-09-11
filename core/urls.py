from django.urls import path
from . import views

urlpatterns = [
    # Public landing page (first screen, before login)
    path('', views.LandingView.as_view(), name='landing'),

    # Auth
    path('login/', views.AdminLoginView.as_view(), name='login'),
    path('logout/', views.AdminLogoutView.as_view(), name='logout'),

    # Dashboard
    path('dashboard/', views.DashboardView.as_view(), name='dashboard'),

    # Dean of Studies Dashboard: oversight of every generated document
    path('dos/dashboard/', views.DoSDashboardView.as_view(), name='dos-dashboard'),
    path(
        'dos/documents/<int:pk>/download/',
        views.GeneratedDocumentDownloadView.as_view(),
        name='dos-document-download'
    ),
    path(
        'dos/documents/<int:pk>/delete/',
        views.GeneratedDocumentDeleteView.as_view(),
        name='dos-document-delete'
    ),

    # Sectors
    path('sectors/', views.SectorListView.as_view(), name='sector-list'),
    path('sectors/create/', views.SectorCreateView.as_view(), name='sector-create'),
    path('sectors/<int:pk>/edit/', views.SectorUpdateView.as_view(), name='sector-edit'),
    path('sectors/<int:pk>/delete/', views.SectorDeleteView.as_view(), name='sector-delete'),

    # Trades
    path('trades/', views.TradeListView.as_view(), name='trade-list'),
    path('trades/create/', views.TradeCreateView.as_view(), name='trade-create'),
    path('trades/<int:pk>/edit/', views.TradeUpdateView.as_view(), name='trade-edit'),
    path('trades/<int:pk>/delete/', views.TradeDeleteView.as_view(), name='trade-delete'),

    # Levels
    path('levels/', views.LevelListView.as_view(), name='level-list'),
    path('levels/create/', views.LevelCreateView.as_view(), name='level-create'),
    path('levels/<int:pk>/edit/', views.LevelUpdateView.as_view(), name='level-edit'),
    path('levels/<int:pk>/delete/', views.LevelDeleteView.as_view(), name='level-delete'),

    # Trade Levels
    path('trade-levels/', views.TradeLevelListView.as_view(), name='tradelevel-list'),
    path('trade-levels/create/', views.TradeLevelCreateView.as_view(), name='tradelevel-create'),
    path('trade-levels/<int:pk>/edit/', views.TradeLevelUpdateView.as_view(), name='tradelevel-edit'),
    path('trade-levels/<int:pk>/delete/', views.TradeLevelDeleteView.as_view(), name='tradelevel-delete'),

    # Trainers
    path('trainers/', views.TrainerListView.as_view(), name='trainer-list'),
    path('trainers/create/', views.TrainerCreateView.as_view(), name='trainer-create'),
    path('trainers/<int:pk>/edit/', views.TrainerUpdateView.as_view(), name='trainer-edit'),
    path('trainers/<int:pk>/delete/', views.TrainerDeleteView.as_view(), name='trainer-delete'),
    path('trainers/<int:pk>/login/toggle/', views.TrainerLoginToggleView.as_view(), name='trainer-login-toggle'),
    path('trainers/<int:pk>/login/repair/', views.TrainerLoginRepairView.as_view(), name='trainer-login-repair'),

    # Trainer Access (grant/revoke generator access from the Dashboard)
    path('trainer-access/', views.TrainerAccessListView.as_view(), name='traineraccess-list'),
    path('trainer-access/<int:pk>/grant/', views.TrainerAccessGrantView.as_view(), name='traineraccess-grant'),
    path('trainer-access/<int:pk>/revoke/', views.TrainerAccessRevokeView.as_view(), name='traineraccess-revoke'),
    path('trainer-access/<int:pk>/edit/', views.TrainerAccessUpdateView.as_view(), name='traineraccess-edit'),

    # Modules
    path('modules/', views.ModuleListView.as_view(), name='module-list'),
    path('modules/create/', views.ModuleCreateView.as_view(), name='module-create'),
    path('modules/<int:pk>/edit/', views.ModuleUpdateView.as_view(), name='module-edit'),
    path('modules/<int:pk>/delete/', views.ModuleDeleteView.as_view(), name='module-delete'),
    path('modules/<int:pk>/toggle/', views.ModuleToggleActiveView.as_view(), name='module-toggle-active'),

    # Learning Outcomes
    path('learning-outcomes/', views.LearningOutcomeListView.as_view(), name='outcome-list'),
    path('learning-outcomes/create/', views.LearningOutcomeCreateView.as_view(), name='outcome-create'),
    path('learning-outcomes/bulk-create/', views.LearningOutcomeBulkCreateView.as_view(), name='outcome-bulk-create'),
    path('learning-outcomes/<int:pk>/edit/', views.LearningOutcomeUpdateView.as_view(), name='outcome-edit'),
    path('learning-outcomes/<int:pk>/delete/', views.LearningOutcomeDeleteView.as_view(), name='outcome-delete'),

    # Indicative Contents
    path('indicative-contents/', views.IndicativeContentListView.as_view(), name='content-list'),
    path('indicative-contents/create/', views.IndicativeContentCreateView.as_view(), name='content-create'),
    path('indicative-contents/bulk-create/', views.IndicativeContentBulkCreateView.as_view(), name='content-bulk-create'),
    path('indicative-contents/<int:pk>/edit/', views.IndicativeContentUpdateView.as_view(), name='content-edit'),
    path('indicative-contents/<int:pk>/delete/', views.IndicativeContentDeleteView.as_view(), name='content-delete'),

    # Logos
    path('logos/', views.LogoListView.as_view(), name='logo-list'),
    path('logos/create/', views.LogoCreateView.as_view(), name='logo-create'),
    path('logos/<int:pk>/edit/', views.LogoUpdateView.as_view(), name='logo-edit'),
    path('logos/<int:pk>/delete/', views.LogoDeleteView.as_view(), name='logo-delete'),

    # Lesson Plans
    path('lesson-plans/', views.LessonPlanListView.as_view(), name='lessonplan-list'),
    path('lesson-plans/create/', views.LessonPlanCreateView.as_view(), name='lessonplan-create'),
    path('lesson-plans/<int:pk>/edit/', views.LessonPlanUpdateView.as_view(), name='lessonplan-edit'),
    path('lesson-plans/<int:pk>/delete/', views.LessonPlanDeleteView.as_view(), name='lessonplan-delete'),

    # Assessment Plans
    path('assessment-plans/', views.AssessmentPlanListView.as_view(), name='assessmentplan-list'),
    path('assessment-plans/create/', views.AssessmentPlanCreateView.as_view(), name='assessmentplan-create'),
    path('assessment-plans/<int:pk>/edit/', views.AssessmentPlanUpdateView.as_view(), name='assessmentplan-edit'),
    path('assessment-plans/<int:pk>/delete/', views.AssessmentPlanDeleteView.as_view(), name='assessmentplan-delete'),

    path('scheme-of-work/user/', views.SchemeOfWorkUserView.as_view(), name='schemeofwork-user'),
    path('lesson-plan/user/', views.LessonPlanUserView.as_view(), name='lessonplan-user'),
    path('assessment-plan/user/', views.AssessmentPlanUserView.as_view(), name='assessmentplan-user'),
    path(
        'api/scheme-of-work/ai-generate/',
        views.generate_scheme_ai_content,
        name='ai-generate'
    ),
    path(
        'api/lesson-plan/ai-generate/',
        views.generate_lesson_plan_ai_content,
        name='lessonplan-ai-generate'
    ),
    path(
        'api/assessment-plan/ai-generate/',
        views.generate_assessment_plan_ai_content,
        name='assessmentplan-ai-generate'
    ),
    path(
        'api/lesson-plan/notes/generate/',
        views.generate_lesson_plan_notes,
        name='lessonplan-notes-generate'
    ),

    # Document export ("Download as Word/Excel") for the three generator pages
    path(
        'api/scheme-of-work/export/docx/',
        views.export_scheme_of_work_docx,
        name='schemeofwork-export-docx'
    ),
    path(
        'api/lesson-plan/export/docx/',
        views.export_lesson_plan_docx,
        name='lessonplan-export-docx'
    ),
    path(
        'api/assessment-plan/export/xlsx/',
        views.export_assessment_plan_xlsx,
        name='assessmentplan-export-xlsx'
    ),
    path(
        'api/scheme-of-work/export/pdf/',
        views.export_scheme_of_work_pdf,
        name='schemeofwork-export-pdf'
    ),
    path(
        'api/lesson-plan/export/pdf/',
        views.export_lesson_plan_pdf,
        name='lessonplan-export-pdf'
    ),
    path(
        'api/assessment-plan/export/pdf/',
        views.export_assessment_plan_pdf,
        name='assessmentplan-export-pdf'
    ),
    path(
        'api/generation/check-access/',
        views.check_generation_access,
        name='generation-check-access'
    ),
]