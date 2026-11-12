/**
 * Protim (TimPRO) — weekly review.
 * Sheet "Protim (Timpro)": daily funnel app DAU → open Protim menu → tap confirm → confirmation succeeds.
 */

function protimFeature_() {
  return createFunnelFeature_({
    id: 'protim',
    name: 'Protim',
    menuFunction: 'generateProtimSlides',
    sheet: 'Protim (Timpro)',
    description: 'Appreciation loyalty program for pengurus: for every fruit (FFB/TBS) delivery they claim the tonnage in the TimPRO ' +
      'app. A claim is only possible after the agent has claimed the same delivery first. ' +
      'Funnel: open the Protim menu → tap confirm → confirmation succeeds.',
    app: 'timpro',
    // Steps 3–4 only have daily distinct users: show the funnel as average users per day, not a weekly sum
    funnelMode: 'daily',
    fields: {
      date: ['date', 'tanggal'],
      dau: ['app_dau', 'daily_active_users'],
      menu: ['step_2_protim_menu_clicked', 'step2_protim_menu_clicked'],
      confirm: ['step3_confirmed_clicked', 'step_3_confirmed_clicked'],
      success: ['step4_confirmed_success', 'step_4_confirmed_success'],
      newUsers: ['new_users_protim'],
      returningUsers: ['returning_users_protim'],
      names: ['user_name_step2', 'user_names_step2'],
      appWau: ['weekly_app_wau', 'app_weekly_active_user'],
      featureWau: ['feature_weekly_wau', 'weekly_feature_wau']
    },
    steps: [
      { key: 'menu', label: 'Open Protim menu', short: 'Menu' },
      // With weekly distinct-user columns for these steps the funnel switches to weekly (app WAU → …)
      { key: 'confirm', label: 'Tap confirm', short: 'Confirm',
        weeklyAliases: ['step3_weekly_users', 'step3_confirmed_weekly_users', 'weekly_step3_users', 'step3_wau'] },
      { key: 'success', label: 'Confirm success', short: 'Success',
        weeklyAliases: ['step4_weekly_users', 'step4_confirmed_success_weekly_users', 'weekly_step4_users', 'step4_wau'] }
    ],
    usersLabel: 'USERS OPEN PROTIM MENU'
  });
}

function generateProtimSlides() {
  generateFeatures_(['protim']);
}
